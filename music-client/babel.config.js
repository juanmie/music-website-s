
module.exports = {
  presets: [
    [
      "@babel/preset-env",
      {
        targets: "> 0.25%, not dead",   // 自动适配主流浏览器最新2个版本‌:ml-citation{ref="3" data="citationList"}
        useBuiltIns: "usage",           // 按需注入polyfill（推荐模式）‌:ml-citation{ref="3,7" data="citationList"}
        corejs: "^3.21.1",               // 指定Core-JS 3最新版‌:ml-citation{ref="3,8" data="citationList"}
        modules: false                  // 保留ES模块语法（优化Tree-Shaking）‌:ml-citation{ref="3" data="citationList"}
      }
    ],
    "@babel/preset-typescript",          // TypeScript支持‌:ml-citation{ref="8" data="citationList"}
    '@vue/cli-plugin-babel/preset'
  ],
  plugins: [
    "@babel/plugin-transform-runtime",  // 复用工具函数（减少代码体积）‌:ml-citation{ref="1,5" data="citationList"}
    "@babel/plugin-proposal-class-properties" // 支持类属性语法‌:ml-citation{ref="4" data="citationList"}
  ]
};