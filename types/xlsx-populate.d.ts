declare module 'xlsx-populate' {
  const XlsxPopulate: {
    fromDataAsync(data: Buffer | ArrayBuffer | Uint8Array, options?: { password?: string }): Promise<{
      outputAsync(type?: string, options?: { password?: string }): Promise<Buffer | ArrayBuffer | Uint8Array>
    }>
  }

  export default XlsxPopulate
}
