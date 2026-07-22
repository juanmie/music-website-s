const { defineConfig } = require('@vue/cli-service')

module.exports = defineConfig({
  transpileDependencies: true,
  devServer: {
    host: '0.0.0.0',
    port: 8080,
    proxy: {
      '/demandEvaluation': {
        target: 'http://localhost:8888',
        changeOrigin: true
      },
      '/ai': {
        target: 'http://localhost:8888',
        changeOrigin: true
      }
    }
  },
  chainWebpack: config => {
    // 允许通过环境变量 NODE_HOST 覆盖（Docker 构建时传入）
    const nodeHost = process.env.NODE_HOST || '"/"';
    config.plugin('define').tap(definitions => {
      Object.assign(definitions[0]['process.env'], {
        NODE_HOST: nodeHost,
      });
      return definitions;
    });
  }
})
