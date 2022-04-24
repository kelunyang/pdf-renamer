const { defineConfig } = require('@vue/cli-service')
module.exports = defineConfig({
  transpileDependencies: [
    'vuetify'
  ],
  pluginOptions: {
    electronBuilder: {
      preload: 'src/preload.js',      
    }
  },
  chainWebpack: config =>{
    config.plugin("copy").use(require('copy-webpack-plugin'), [{
      patterns: [
        {
          from: 'node_modules/pdfjs-dist/build/pdf.worker.js',
          to: 'pdf.worker.js'
       }
      ]
    }])
  }
})
