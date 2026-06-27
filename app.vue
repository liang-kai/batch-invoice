<script setup lang="ts">
type SummaryItem = {
  rowNumber: number
  orderId: string
  recipient: string
  trackingNo?: string
  logisticsCompany?: string
  shipmentRecipient?: string
  reason?: string
}

type FillResponse = {
  fileName: string
  mimeType: string
  base64: string
  summary: {
    totalOrders: number
    totalShipments: number
    successCount: number
    failedCount: number
    usedModel: boolean
    modelName: string
    successes: SummaryItem[]
    failures: SummaryItem[]
  }
}

const orderFile = ref<File | null>(null)
const shipmentFile = ref<File | null>(null)
const apiKey = ref('')
const baseUrl = ref('https://api.openai.com/v1')
const model = ref('gpt-4.1-mini')
const loading = ref(false)
const errorMessage = ref('')
const result = ref<FillResponse | null>(null)

const canSubmit = computed(() => Boolean(orderFile.value && shipmentFile.value && !loading.value))

function onFileChange(event: Event, kind: 'order' | 'shipment') {
  const files = (event.target as HTMLInputElement).files
  const file = files?.[0] ?? null
  if (kind === 'order') orderFile.value = file
  if (kind === 'shipment') shipmentFile.value = file
  result.value = null
  errorMessage.value = ''
}

function downloadResult() {
  if (!result.value) return
  const binary = atob(result.value.base64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  const blob = new Blob([bytes], { type: result.value.mimeType })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = result.value.fileName
  link.click()
  URL.revokeObjectURL(url)
}

async function submit() {
  if (!orderFile.value || !shipmentFile.value) return
  loading.value = true
  errorMessage.value = ''
  result.value = null

  try {
    const form = new FormData()
    form.append('orderFile', orderFile.value)
    form.append('shipmentFile', shipmentFile.value)
    form.append('apiKey', apiKey.value)
    form.append('baseUrl', baseUrl.value)
    form.append('model', model.value)

    result.value = await $fetch<FillResponse>('/api/fill', {
      method: 'POST',
      body: form,
    })
  } catch (error: any) {
    errorMessage.value = error?.data?.message || error?.message || '处理失败，请检查文件格式后重试。'
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <main class="page">
    <section class="workspace">
      <header class="topbar">
        <div>
          <p class="eyebrow">快团团发货单助手</p>
          <h1>上传未发货订单和快递单，一键回填物流信息</h1>
        </div>
        <button class="download" :disabled="!result" @click="downloadResult">下载发货单</button>
      </header>

      <form class="panel" @submit.prevent="submit">
        <div class="upload-grid">
          <label class="drop">
            <span class="label">1. 快团团未发货订单</span>
            <strong>{{ orderFile?.name || '选择 0626单子.xlsx' }}</strong>
            <input type="file" accept=".xlsx,.xls" @change="onFileChange($event, 'order')" />
          </label>

          <label class="drop">
            <span class="label">2. 快递公司订单 Excel</span>
            <strong>{{ shipmentFile?.name || '选择自建打单文件' }}</strong>
            <input type="file" accept=".xlsx,.xls" @change="onFileChange($event, 'shipment')" />
          </label>
        </div>

        <div class="settings">
          <label>
            <span>OpenAI 兼容 API Key</span>
            <input v-model="apiKey" type="password" placeholder="留空则使用服务端环境变量；都没有时使用本地规则匹配" autocomplete="off" />
          </label>
          <label>
            <span>Base URL</span>
            <input v-model="baseUrl" type="url" placeholder="https://api.openai.com/v1" />
          </label>
          <label>
            <span>模型</span>
            <input v-model="model" type="text" placeholder="gpt-4.1-mini" />
          </label>
        </div>

        <div class="actions">
          <button class="primary" :disabled="!canSubmit" type="submit">
            {{ loading ? '正在解析并填充...' : '生成发货单' }}
          </button>
          <p v-if="errorMessage" class="error">{{ errorMessage }}</p>
        </div>
      </form>

      <section v-if="result" class="results">
        <div class="metrics">
          <div>
            <span>未发货订单</span>
            <strong>{{ result.summary.totalOrders }}</strong>
          </div>
          <div>
            <span>快递记录</span>
            <strong>{{ result.summary.totalShipments }}</strong>
          </div>
          <div>
            <span>成功</span>
            <strong>{{ result.summary.successCount }}</strong>
          </div>
          <div>
            <span>失败</span>
            <strong>{{ result.summary.failedCount }}</strong>
          </div>
        </div>

        <div class="status-line">
          <span :class="['pill', result.summary.usedModel ? 'ok' : 'warn']">
            {{ result.summary.usedModel ? `已调用模型：${result.summary.modelName}` : '未调用模型，使用本地规则匹配' }}
          </span>
          <button class="secondary" @click="downloadResult">下载 {{ result.fileName }}</button>
        </div>

        <div class="tables">
          <article>
            <h2>成功匹配</h2>
            <table>
              <thead>
                <tr>
                  <th>行号</th>
                  <th>订单号</th>
                  <th>收货人</th>
                  <th>运单号</th>
                  <th>物流公司</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="item in result.summary.successes" :key="`${item.rowNumber}-${item.trackingNo}`">
                  <td>{{ item.rowNumber }}</td>
                  <td>{{ item.orderId }}</td>
                  <td>{{ item.recipient }}</td>
                  <td>{{ item.trackingNo }}</td>
                  <td>{{ item.logisticsCompany }}</td>
                </tr>
              </tbody>
            </table>
          </article>

          <article>
            <h2>需要人工确认</h2>
            <table>
              <thead>
                <tr>
                  <th>行号</th>
                  <th>订单号</th>
                  <th>收货人</th>
                  <th>原因</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="item in result.summary.failures" :key="`${item.rowNumber}-${item.orderId}`">
                  <td>{{ item.rowNumber }}</td>
                  <td>{{ item.orderId }}</td>
                  <td>{{ item.recipient }}</td>
                  <td>{{ item.reason }}</td>
                </tr>
                <tr v-if="!result.summary.failures.length">
                  <td colspan="4">全部订单都完成了。</td>
                </tr>
              </tbody>
            </table>
          </article>
        </div>
      </section>
    </section>
  </main>
</template>

<style scoped>
.page {
  min-height: 100vh;
  padding: 32px;
}

.workspace {
  width: min(1180px, 100%);
  margin: 0 auto;
}

.topbar {
  display: flex;
  align-items: end;
  justify-content: space-between;
  gap: 24px;
  margin-bottom: 20px;
}

.eyebrow {
  margin: 0 0 8px;
  color: #59735f;
  font-size: 14px;
  font-weight: 700;
}

h1 {
  margin: 0;
  max-width: 760px;
  font-size: clamp(30px, 4vw, 56px);
  line-height: 1.04;
  letter-spacing: 0;
}

.panel,
.results {
  border: 1px solid #d9d7c9;
  background: #fffdf6;
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 18px 44px rgb(42 52 34 / 10%);
}

.upload-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 16px;
}

.drop {
  display: grid;
  gap: 10px;
  min-height: 150px;
  padding: 20px;
  border: 1px dashed #9aa88f;
  border-radius: 8px;
  background: #f7f9f2;
}

.drop strong {
  overflow-wrap: anywhere;
  font-size: 18px;
}

.label,
.settings span,
.metrics span {
  color: #607064;
  font-size: 13px;
  font-weight: 700;
}

.settings {
  display: grid;
  grid-template-columns: 1.3fr 1fr 0.8fr;
  gap: 14px;
  margin-top: 18px;
}

.settings label {
  display: grid;
  gap: 8px;
}

.settings input {
  width: 100%;
  min-height: 42px;
  border: 1px solid #cfcdbf;
  border-radius: 6px;
  padding: 0 12px;
  background: white;
  color: #17211b;
}

.actions,
.status-line {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  margin-top: 18px;
}

button {
  min-height: 42px;
  border: 0;
  border-radius: 6px;
  padding: 0 16px;
  font-weight: 800;
}

.primary,
.download {
  background: #1d5836;
  color: white;
}

.secondary {
  background: #e6efe4;
  color: #1d5836;
}

button:disabled {
  opacity: 0.48;
}

.error {
  margin: 0;
  color: #a43c24;
  font-weight: 700;
}

.results {
  margin-top: 20px;
}

.metrics {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 14px;
}

.metrics div {
  display: grid;
  gap: 8px;
  padding: 16px;
  border-radius: 8px;
  background: #f2f5ec;
}

.metrics strong {
  font-size: 32px;
}

.pill {
  display: inline-flex;
  align-items: center;
  min-height: 32px;
  border-radius: 999px;
  padding: 0 12px;
  font-weight: 800;
}

.pill.ok {
  background: #dceee2;
  color: #1d5836;
}

.pill.warn {
  background: #fff0cc;
  color: #7a5417;
}

.tables {
  display: grid;
  grid-template-columns: 1.2fr 1fr;
  gap: 18px;
  margin-top: 18px;
}

article {
  min-width: 0;
}

h2 {
  margin: 0 0 10px;
  font-size: 18px;
}

table {
  width: 100%;
  border-collapse: collapse;
  overflow: hidden;
  border-radius: 8px;
  font-size: 13px;
}

th,
td {
  max-width: 240px;
  padding: 10px;
  border-bottom: 1px solid #e6e2d5;
  text-align: left;
  vertical-align: top;
  overflow-wrap: anywhere;
}

th {
  background: #eef1e8;
  color: #46574a;
}

@media (max-width: 880px) {
  .page {
    padding: 18px;
  }

  .topbar,
  .actions,
  .status-line {
    align-items: stretch;
    flex-direction: column;
  }

  .upload-grid,
  .settings,
  .metrics,
  .tables {
    grid-template-columns: 1fr;
  }
}
</style>
