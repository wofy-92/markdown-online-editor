<!-- @format -->

<template>
  <div class="export-page">
    <div class="button-group">
      <el-button round @click="onBackToMainPage">返回主页</el-button>
      <el-button round @click="onExportBtnClick" type="primary" :disabled="exporting">
        {{ exporting ? '正在导出...' : '生成导出' }}
      </el-button>
    </div>
    <PreviewVditor :pdata="pdata" />
  </div>
</template>

<script>
import PreviewVditor from '@components/PreviewVditor'
import { getExportFileName } from '@helper/utils'
import { getActiveDocId, getDocContent } from '@helper/storage'
import { trackEvent } from '@helper/analytics'

export default {
  name: 'export-docx',

  data() {
    return {
      isLoading: true,
      pdata: getDocContent(getActiveDocId()) || '',
      exporting: false,
    }
  },

  components: {
    PreviewVditor,
  },

  methods: {
    loadHTMLtoDOCX() {
      return new Promise((resolve, reject) => {
        if (window.HTMLToDOCX) {
          resolve(window.HTMLToDOCX)
          return
        }
        // Polyfill for Node.js global object in browser
        if (typeof window.global === 'undefined') {
          window.global = window
        }
        const script = document.createElement('script')
        script.src = '/vendor/html-to-docx.browser.js'
        script.onload = () => resolve(window.HTMLToDOCX)
        script.onerror = () => reject(new Error('Failed to load html-to-docx'))
        document.head.appendChild(script)
      })
    },

    async exportAndDownloadDocx(contentElement, filename) {
      try {
        const HTMLtoDOCX = await this.loadHTMLtoDOCX()
        const htmlContent = contentElement.innerHTML

        const fullHtml = `
          <!DOCTYPE html>
          <html>
            <head>
              <meta charset="UTF-8">
              <style>
                body { font-family: 'Microsoft YaHei', 'SimSun', sans-serif; }
                img { max-width: 100%; height: auto; }
                table { border-collapse: collapse; width: 100%; }
                td, th { border: 1px solid #ccc; padding: 8px; }
              </style>
            </head>
            <body>
              ${htmlContent}
            </body>
          </html>
        `

        const docxBlob = await HTMLtoDOCX(fullHtml, null, {
          table: { row: { cantSplit: true } },
          footer: true,
          pageNumber: true,
        })

        const url = URL.createObjectURL(docxBlob)
        const link = document.createElement('a')
        link.href = url
        link.download = filename
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
        URL.revokeObjectURL(url)

        this.$message.success('Word 导出成功')
      } catch (error) {
        console.error('Word 导出失败:', error)
        this.$message.error('Word 导出失败，请重试')
      } finally {
        this.isLoading = false
        this.exporting = false
      }
    },

    onBackToMainPage() {
      this.$router.push('/')
    },

    onExportBtnClick() {
      this.isLoading = true
      this.exporting = true
      const contentElement = document.querySelector('#khaleesi .vditor-preview .vditor-reset')
      const filename = getExportFileName()
      this.exportAndDownloadDocx(contentElement, filename)
      trackEvent('export_docx_submit', 'export', filename)
    },
  },
}
</script>
