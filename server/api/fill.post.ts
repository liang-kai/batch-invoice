import ExcelJS from 'exceljs'
import { createError, readMultipartFormData } from 'h3'

type ParsedRow = {
  id: string
  rowNumber: number
  values: Record<string, string>
}

type Candidate = {
  shipmentId: string
  score: number
  reasons: string[]
}

type ModelMatch = {
  orderId: string
  shipmentId?: string | null
  shipmentIds?: string[] | null
  confidence?: number
  reason?: string
}

type FinalMatch = {
  order: ParsedRow
  shipments?: ParsedRow[]
  reason: string
}

const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function cellToString(value: ExcelJS.CellValue): string {
  if (value == null) return ''
  if (value instanceof Date) return value.toISOString().replace('T', ' ').slice(0, 19)
  if (typeof value === 'object') {
    if ('text' in value && typeof value.text === 'string') return value.text
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text ?? '').join('')
    }
    if ('result' in value) return cellToString(value.result as ExcelJS.CellValue)
    if ('hyperlink' in value && 'text' in value) return String(value.text ?? value.hyperlink ?? '')
  }
  return String(value).trim()
}

function normalize(value: string): string {
  return value
    .replace(/\s+/g, '')
    .replace(/[，,。.\-—_()（）[\]【】]/g, '')
    .toLowerCase()
}

function onlyDigits(value: string): string {
  return value.replace(/\D/g, '')
}

function findHeaderIndex(headers: string[], names: string[]): number {
  return headers.findIndex((header) => names.some((name) => header.includes(name)))
}

function get(row: ParsedRow, names: string[]): string {
  const entry = Object.entries(row.values).find(([key]) => names.some((name) => key.includes(name)))
  return entry?.[1] ?? ''
}

async function parseWorkbook(buffer: Buffer) {
  const workbook = new ExcelJS.Workbook()
  const source = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
  await workbook.xlsx.load(source)
  const worksheet = workbook.worksheets[0]
  if (!worksheet) {
    throw createError({ statusCode: 400, message: 'Excel 文件里没有工作表。' })
  }

  const headerRow = worksheet.getRow(1)
  const headers = Array.from({ length: worksheet.columnCount }, (_, index) => {
    return cellToString(headerRow.getCell(index + 1).value) || `列${index + 1}`
  })

  const rows: ParsedRow[] = []
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber)
    const values: Record<string, string> = {}
    headers.forEach((header, index) => {
      values[header] = cellToString(row.getCell(index + 1).value)
    })
    if (Object.values(values).some(Boolean)) {
      rows.push({ id: String(rowNumber), rowNumber, values })
    }
  }

  return { workbook, worksheet, headers, rows }
}

function scoreCandidate(order: ParsedRow, shipment: ParsedRow): Candidate {
  const reasons: string[] = []
  let score = 0

  const orderName = normalize(get(order, ['收货人', '收件人', '下单人']).replace(/\[\d+\]/g, ''))
  const shipmentName = normalize(get(shipment, ['收件人姓名', '收货人', '收件人']))
  if (orderName && shipmentName) {
    if (orderName === shipmentName) {
      score += 45
      reasons.push('姓名一致')
    } else if (orderName.includes(shipmentName) || shipmentName.includes(orderName)) {
      score += 28
      reasons.push('姓名相近')
    }
  }

  for (const field of ['省', '市', '区']) {
    const orderPart = normalize(get(order, [field]))
    const shipmentPart = normalize(get(shipment, [`收件人${field}`, field]))
    if (orderPart && shipmentPart && orderPart === shipmentPart) {
      score += 10
      reasons.push(`${field}一致`)
    }
  }

  const orderAddress = normalize(get(order, ['详细地址', '地址']))
  const shipmentAddress = normalize(get(shipment, ['收件人详细地址', '详细地址', '地址']))
  if (orderAddress && shipmentAddress) {
    if (orderAddress.includes(shipmentAddress) || shipmentAddress.includes(orderAddress)) {
      score += 60
      reasons.push('地址包含匹配')
    } else {
      const sharedTokens = shipmentAddress
        .match(/[\u4e00-\u9fa5a-z0-9]{2,}/g)
        ?.filter((token) => token.length >= 3 && orderAddress.includes(token))
      if (sharedTokens?.length) {
        score += Math.min(35, sharedTokens.length * 7)
        reasons.push('地址片段相同')
      }
    }
  }

  const orderPhoneDigits = onlyDigits(`${get(order, ['联系电话', '手机'])} ${get(order, ['详细地址'])}`)
  const shipmentPhoneDigits = onlyDigits(get(shipment, ['收件人手机/电话', '联系电话', '手机']))
  if (orderPhoneDigits && shipmentPhoneDigits) {
    if (orderPhoneDigits.includes(shipmentPhoneDigits) || shipmentPhoneDigits.includes(orderPhoneDigits)) {
      score += 55
      reasons.push('手机号一致')
    } else if (orderPhoneDigits.slice(-4) === shipmentPhoneDigits.slice(-4)) {
      score += 26
      reasons.push('手机号尾号一致')
    }
  }

  return { shipmentId: shipment.id, score, reasons }
}

function buildCandidates(orders: ParsedRow[], shipments: ParsedRow[]) {
  return orders.map((order) => {
    const candidates = shipments
      .map((shipment) => scoreCandidate(order, shipment))
      .sort((a, b) => b.score - a.score)
      .slice(0, 8)
    return { order, candidates }
  })
}

function localMatches(orders: ParsedRow[], shipments: ParsedRow[]): ModelMatch[] {
  const used = new Set<string>()
  return buildCandidates(orders, shipments).map(({ order, candidates }) => {
    const best = candidates.find((candidate) => !used.has(candidate.shipmentId))
    const second = candidates.find((candidate) => candidate.shipmentId !== best?.shipmentId)
    const tied = best
      ? candidates.filter((candidate) => {
          return !used.has(candidate.shipmentId) && candidate.score >= 60 && Math.abs(candidate.score - best.score) <= 3
        })
      : []
    const confident = best && best.score >= 60 && (best.score - (second?.score ?? 0) >= 10 || tied.length > 1)
    if (confident) {
      for (const candidate of tied.length > 1 ? tied : [best]) used.add(candidate.shipmentId)
      return {
        orderId: order.id,
        shipmentIds: (tied.length > 1 ? tied : [best]).map((candidate) => candidate.shipmentId),
        confidence: Math.min(0.95, best.score / 120),
        reason: tied.length > 1 ? `多箱匹配：${tied.length} 个运单` : best.reasons.join('、') || '本地规则匹配',
      }
    }
    return {
      orderId: order.id,
      shipmentId: null,
      confidence: 0,
      reason: best ? `候选分数不足：${best.score}` : '没有候选快递记录',
    }
  })
}

function compactOrder(row: ParsedRow) {
  return {
    orderId: row.id,
    rowNumber: row.rowNumber,
    kttOrderNo: get(row, ['订单号']),
    recipient: get(row, ['收货人']),
    phone: get(row, ['联系电话', '手机']),
    province: get(row, ['省']),
    city: get(row, ['市']),
    district: get(row, ['区']),
    address: get(row, ['详细地址']),
    buyer: get(row, ['下单人']),
  }
}

function compactShipment(row: ParsedRow) {
  return {
    shipmentId: row.id,
    recipient: get(row, ['收件人姓名', '收货人']),
    phone: get(row, ['收件人手机/电话', '联系电话', '手机']),
    province: get(row, ['收件人省', '省']),
    city: get(row, ['收件人市', '市']),
    district: get(row, ['收件人区', '区']),
    address: get(row, ['收件人详细地址', '详细地址']),
    trackingNo: get(row, ['运单号', '物流单号', '快递单号']),
  }
}

function extractJsonArray(text: string): ModelMatch[] {
  const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1]
  const raw = fenced || text
  const start = raw.indexOf('[')
  const end = raw.lastIndexOf(']')
  if (start === -1 || end === -1 || end <= start) {
    throw new Error('模型没有返回 JSON 数组')
  }
  return JSON.parse(raw.slice(start, end + 1))
}

async function callOpenAICompatible(input: {
  apiKey: string
  baseUrl: string
  model: string
  orders: ParsedRow[]
  shipments: ParsedRow[]
}) {
  const candidateGroups = buildCandidates(input.orders, input.shipments)
  const matches: ModelMatch[] = []

  for (let index = 0; index < candidateGroups.length; index += 30) {
    const chunk = candidateGroups.slice(index, index + 30)
    const payload = chunk.map(({ order, candidates }) => ({
      order: compactOrder(order),
      candidates: candidates.map((candidate) => {
        const shipment = input.shipments.find((item) => item.id === candidate.shipmentId)!
        return {
          ...compactShipment(shipment),
          localScore: candidate.score,
          localReasons: candidate.reasons,
        }
      }),
    }))

    const baseUrl = input.baseUrl.replace(/\/$/, '')
    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${input.apiKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: input.model,
        temperature: baseUrl.includes('api.kimi.com/coding') || input.model === 'kimi-for-coding' ? 1 : 0,
        messages: [
          {
            role: 'system',
            content:
              '你是订单发货匹配助手。根据快团团订单和候选快递记录，判断每个订单唯一对应的快递记录。只返回 JSON 数组，不要解释。',
          },
          {
            role: 'user',
            content: [
              '匹配原则：优先收货人、手机号、详细地址、省市区一致；快团团地址里可能带尾号标记，姓名可能带方括号；不确定就返回 null。',
              '返回格式：[{ "orderId": "2", "shipmentIds": ["5"] 或 ["5","6"] 或 null, "confidence": 0-1, "reason": "简短原因" }]',
              '如果一个订单对应多个同地址同收件人的包裹，请把多个 shipmentId 都放入 shipmentIds。',
              JSON.stringify(payload, null, 2),
            ].join('\n\n'),
          },
        ],
      }),
    })

    if (!response.ok) {
      const errorText = await response.text()
      throw new Error(`模型接口调用失败：${response.status} ${errorText.slice(0, 500)}`)
    }

    const data = await response.json()
    const content = data?.choices?.[0]?.message?.content
    if (!content) throw new Error('模型接口没有返回 message.content')
    matches.push(...extractJsonArray(content))
  }

  return matches
}

function detectLogisticsCompany(trackingNo: string): string {
  const value = trackingNo.trim().toUpperCase()
  if (!value) return ''
  if (/^SF\d{10,15}$/.test(value) || /^9\d{11}$/.test(value)) return '顺丰速运'
  if (/^YT\d{13,18}$/.test(value)) return '圆通速递'
  if (/^JD[A-Z0-9]+$/.test(value) || /^JDV[A-Z0-9]+$/.test(value)) return '京东物流'
  if (/^JT\d{12,18}$/.test(value)) return '极兔速递'
  if (/^DPK\d+/.test(value)) return '德邦快递'
  if (/^77\d{13,16}$/.test(value) || /^7\d{14,17}$/.test(value)) return '中通快递'
  if (/^4\d{12,15}$/.test(value)) return '韵达快递'
  if (/^3\d{12,15}$/.test(value)) return '申通快递'
  return '未知快递'
}

function cloneStyle<T>(style: T): T {
  if (!style || typeof style !== 'object') return style
  return JSON.parse(JSON.stringify(style))
}

function cloneCellValue(value: ExcelJS.CellValue): ExcelJS.CellValue {
  if (!value || typeof value !== 'object' || value instanceof Date) return value
  return JSON.parse(JSON.stringify(value))
}

function duplicateOrderRow(worksheet: ExcelJS.Worksheet, sourceRowNumber: number, targetRowNumber: number) {
  const sourceRow = worksheet.getRow(sourceRowNumber)
  worksheet.insertRow(targetRowNumber, [])
  const targetRow = worksheet.getRow(targetRowNumber)
  targetRow.height = sourceRow.height
  targetRow.hidden = sourceRow.hidden
  targetRow.outlineLevel = sourceRow.outlineLevel

  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
    const sourceCell = sourceRow.getCell(columnNumber)
    const targetCell = targetRow.getCell(columnNumber)
    targetCell.value = cloneCellValue(sourceCell.value)
    targetCell.style = cloneStyle(sourceCell.style)
    targetCell.numFmt = sourceCell.numFmt
    targetCell.alignment = cloneStyle(sourceCell.alignment)
    targetCell.border = cloneStyle(sourceCell.border)
    targetCell.fill = cloneStyle(sourceCell.fill)
    targetCell.font = cloneStyle(sourceCell.font)
    targetCell.protection = cloneStyle(sourceCell.protection)
  }
}

function fillShipmentCells(
  worksheet: ExcelJS.Worksheet,
  rowNumber: number,
  companyColumn: number,
  trackingColumn: number,
  trackingNo: string,
) {
  worksheet.getRow(rowNumber).getCell(companyColumn + 1).value = detectLogisticsCompany(trackingNo)
  worksheet.getRow(rowNumber).getCell(trackingColumn + 1).value = trackingNo
}

function makeFinalMatches(orders: ParsedRow[], shipments: ParsedRow[], modelMatches: ModelMatch[]): FinalMatch[] {
  const shipmentById = new Map(shipments.map((shipment) => [shipment.id, shipment]))
  const used = new Set<string>()
  const modelByOrder = new Map(modelMatches.map((match) => [match.orderId, match]))

  return orders.map((order) => {
    const match = modelByOrder.get(order.id)
    const shipmentIds = match?.shipmentIds?.length ? match.shipmentIds : match?.shipmentId ? [match.shipmentId] : []
    if (!shipmentIds.length) {
      return { order, reason: match?.reason || '未匹配到快递记录' }
    }
    const repeated = shipmentIds.find((shipmentId) => used.has(shipmentId))
    if (repeated) {
      return { order, reason: `快递记录 ${repeated} 已被其他订单使用` }
    }
    const matchedShipments = shipmentIds.map((shipmentId) => shipmentById.get(shipmentId))
    if (matchedShipments.some((shipment) => !shipment)) {
      return { order, reason: `模型返回的快递记录不存在：${shipmentIds.join(', ')}` }
    }
    for (const shipmentId of shipmentIds) used.add(shipmentId)
    return { order, shipments: matchedShipments as ParsedRow[], reason: match?.reason || '匹配成功' }
  })
}

export default defineEventHandler(async (event) => {
  const parts = await readMultipartFormData(event)
  if (!parts) {
    throw createError({ statusCode: 400, message: '没有收到上传文件。' })
  }

  const filePart = (name: string) => parts.find((part) => part.name === name && part.data)
  const orderFile = filePart('orderFile')
  const shipmentFile = filePart('shipmentFile')
  if (!orderFile?.data || !shipmentFile?.data) {
    throw createError({ statusCode: 400, message: '请同时上传快团团订单和快递公司订单 Excel。' })
  }

  const parsedOrder = await parseWorkbook(orderFile.data)
  const parsedShipment = await parseWorkbook(shipmentFile.data)

  const companyColumn = findHeaderIndex(parsedOrder.headers, ['物流公司'])
  const trackingColumn = findHeaderIndex(parsedOrder.headers, ['物流单号', '运单号', '快递单号'])
  if (companyColumn === -1 || trackingColumn === -1) {
    throw createError({ statusCode: 400, message: '快团团模板缺少“物流公司”或“物流单号”列。' })
  }

  const apiKey = process.env.OPENAI_API_KEY || ''
  const baseUrl = process.env.OPENAI_BASE_URL || 'https://api.openai.com/v1'
  const model = process.env.OPENAI_MODEL || 'gpt-4.1-mini'

  let usedModel = false
  let matches: ModelMatch[]
  if (apiKey) {
    try {
      matches = await callOpenAICompatible({
        apiKey,
        baseUrl,
        model,
        orders: parsedOrder.rows,
        shipments: parsedShipment.rows,
      })
      usedModel = true
    } catch (error: any) {
      throw createError({
        statusCode: 502,
        message: error?.message || '模型接口调用失败。',
      })
    }
  } else {
    matches = localMatches(parsedOrder.rows, parsedShipment.rows)
  }

  const finalMatches = makeFinalMatches(parsedOrder.rows, parsedShipment.rows, matches)
  const successes: any[] = []
  const failures: any[] = []

  const orderedMatches = [...finalMatches].sort((a, b) => b.order.rowNumber - a.order.rowNumber)

  for (const item of orderedMatches) {
    const orderNo = get(item.order, ['订单号'])
    const recipient = get(item.order, ['收货人'])
    if (!item.shipments?.length) {
      failures.push({
        rowNumber: item.order.rowNumber,
        orderId: orderNo,
        recipient,
        reason: item.reason,
      })
      continue
    }

    const trackingNumbers = item.shipments.map((shipment) => get(shipment, ['运单号', '物流单号', '快递单号'])).filter(Boolean)
    const firstTrackingNo = trackingNumbers[0] ?? ''
    fillShipmentCells(parsedOrder.worksheet, item.order.rowNumber, companyColumn, trackingColumn, firstTrackingNo)

    for (let index = 1; index < trackingNumbers.length; index += 1) {
      const targetRowNumber = item.order.rowNumber + index
      duplicateOrderRow(parsedOrder.worksheet, item.order.rowNumber, targetRowNumber)
      fillShipmentCells(parsedOrder.worksheet, targetRowNumber, companyColumn, trackingColumn, trackingNumbers[index]!)
    }

    successes.push({
      rowNumber: item.order.rowNumber,
      orderId: orderNo,
      recipient,
      trackingNo: trackingNumbers.join(','),
      logisticsCompany: Array.from(new Set(trackingNumbers.map(detectLogisticsCompany).filter(Boolean))).join(',') || '未知快递',
      shipmentRecipient: get(item.shipments[0]!, ['收件人姓名', '收货人']),
      reason: item.reason,
    })
  }

  successes.sort((a, b) => a.rowNumber - b.rowNumber)
  failures.sort((a, b) => a.rowNumber - b.rowNumber)

  const buffer = await parsedOrder.workbook.xlsx.writeBuffer()
  const date = new Date().toISOString().slice(0, 10).replace(/-/g, '')

  return {
    fileName: `快团团发货单-已填充-${date}.xlsx`,
    mimeType: MIME_XLSX,
    base64: Buffer.from(buffer).toString('base64'),
    summary: {
      totalOrders: parsedOrder.rows.length,
      totalShipments: parsedShipment.rows.length,
      successCount: successes.length,
      failedCount: failures.length,
      usedModel,
      modelName: usedModel ? model : 'local-rules',
      successes,
      failures,
    },
  }
})
