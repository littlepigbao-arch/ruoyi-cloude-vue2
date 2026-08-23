'use strict'
const path = require('path')
const fs = require('fs')

function resolve(dir) {
  return path.join(__dirname, dir)
}

const CompressionPlugin = require('compression-webpack-plugin')

const name = process.env.VUE_APP_TITLE || '若依管理系统' // 网页标题

const port = process.env.port || process.env.npm_config_port || 80 // 端口

// ===== Univer 兼容性辅助：动态生成别名 =====
// 背景：Webpack 4 (Vue CLI 4) 不支持 package.json 的 exports 字段，
// 而 @univerjs 各子包、@radix-ui/*、cjk-regex 等纯 ESM 包依赖 exports
// 字段暴露子路径/入口，导致构建期 "Module not found"。这里通过 fs 扫描
// node_modules，自动为 locale/locales 子路径和 @radix-ui 子路径生成别名。
function buildUniverAliases() {
  const alias = {}
  // 1) @univerjs 子包的 locale/locales 子路径 (zh-CN)
  const univerRoot = resolve('node_modules/@univerjs')
  if (fs.existsSync(univerRoot)) {
    fs.readdirSync(univerRoot, { withFileTypes: true }).forEach(entry => {
      if (!entry.isDirectory()) return
      const pkg = entry.name
      const localeSingle = resolve(`node_modules/@univerjs/${pkg}/lib/cjs/locale/zh-CN.js`)
      if (fs.existsSync(localeSingle)) {
        alias[`@univerjs/${pkg}/locale/zh-CN`] = localeSingle
      }
      const localePlural = resolve(`node_modules/@univerjs/${pkg}/lib/cjs/locales/zh-CN.js`)
      if (fs.existsSync(localePlural)) {
        alias[`@univerjs/${pkg}/locales/zh-CN`] = localePlural
      }
    })
  }
  // 2) @radix-ui/* 子路径 (Webpack 4 无法解析其 exports 条件字段)
  const radixRoot = resolve('node_modules/@radix-ui')
  if (fs.existsSync(radixRoot)) {
    fs.readdirSync(radixRoot, { withFileTypes: true }).forEach(entry => {
      if (!entry.isDirectory()) return
      const pkg = entry.name
      // 2.1 @radix-ui/<pkg>/is-development 子路径
      const isDevCjs = resolve(`node_modules/@radix-ui/${pkg}/dist/internal/is-development.${process.env.NODE_ENV === 'production' ? 'false' : 'true'}.js`)
      const isDevCjsFallback = resolve(`node_modules/@radix-ui/${pkg}/dist/internal/is-development.false.js`)
      if (fs.existsSync(isDevCjs)) {
        alias[`@radix-ui/${pkg}/is-development`] = isDevCjs
      } else if (fs.existsSync(isDevCjsFallback)) {
        alias[`@radix-ui/${pkg}/is-development`] = isDevCjsFallback
      }
    })
  }
  return alias
}

const univerAliases = buildUniverAliases()

// vue.config.js 配置说明
//官方vue.config.js 参考文档 https://cli.vuejs.org/zh/config/#css-loaderoptions
// 这里只列一部分，具体配置参考文档
module.exports = {
  // 部署生产环境和开发环境下的URL。
  // 默认情况下，Vue CLI 会假设你的应用是被部署在一个域名的根路径上
  // 例如 https://www.ruoyi.vip/。如果应用被部署在一个子路径上，你就需要用这个选项指定这个子路径。例如，如果你的应用被部署在 https://www.ruoyi.vip/admin/，则设置 baseUrl 为 /admin/。
  publicPath: process.env.NODE_ENV === "production" ? "/" : "/",
  // 在npm run build 或 yarn build 时 ，生成文件的目录名称（要和baseUrl的生产环境路径一致）（默认dist）
  outputDir: 'dist',
  // 用于放置生成的静态资源 (js、css、img、fonts) 的；（项目打包之后，静态资源会放在这个文件夹下）
  assetsDir: 'static',
  // 如果你不需要生产环境的 source map，可以将其设置为 false 以加速生产环境构建。
  productionSourceMap: false,
  transpileDependencies: [
    'quill',
    // ===== Univer 相关依赖需要转译 ES6+ 新语法 (??, class fields, etc.) =====
    '@univerjs',
    '@wendellhu',
    'numfmt',
    'cjk-regex',
    'regexp-util',
    'unicode-regex',
    'lodash-es',
    // Univer UI 依赖的 React 生态 ESM 包 (含 .mjs 文件 + class fields 语法)
    '@radix-ui',
    'class-variance-authority',
    'tailwind-merge'
  ],
  // webpack-dev-server 相关配置
  devServer: {
    host: '0.0.0.0',
    port: port,
    open: true,
    proxy: {
      // detail: https://cli.vuejs.org/config/#devserver-proxy
      [process.env.VUE_APP_BASE_API]: {
        target: `http://localhost:8080`,
        changeOrigin: true,
        pathRewrite: {
          ['^' + process.env.VUE_APP_BASE_API]: ''
        }
      }
    },
    disableHostCheck: true
  },
  css: {
    loaderOptions: {
      sass: {
        sassOptions: { outputStyle: "expanded" }
      }
    }
  },
  configureWebpack: {
    name: name,
    resolve: {
      alias: Object.assign({
        '@': resolve('src'),
        // ===== Univer 兼容性别名 (解决 Webpack 4 不支持 package exports 字段) =====
        // 关键：facade 子路径必须指向 ES 根文件 lib/facade.js，与 @univerjs/presets
        // 实际 import 的 "@univerjs/core/lib/facade" 解析到同一个文件。否则会产生
        // 两个不同的 FUniver 类：presets 用的是 ES 根 FUniver，而 sheets facade
        // 通过 alias 拿到的是 CJS FUniver，导致 FUniver.extend(FUniverSheetsMixin)
        // 扩展了错误的类，univerAPI.getActiveWorkbook() 恒为 undefined。
        '@univerjs/core/facade$': resolve('node_modules/@univerjs/core/lib/facade.js'),
        '@univerjs/sheets/facade$': resolve('node_modules/@univerjs/sheets/lib/facade.js'),
        '@univerjs/sheets-ui/facade$': resolve('node_modules/@univerjs/sheets-ui/lib/facade.js'),
        '@univerjs/engine-formula/facade$': resolve('node_modules/@univerjs/engine-formula/lib/facade.js'),
        '@univerjs/network/facade$': resolve('node_modules/@univerjs/network/lib/facade.js'),
        '@univerjs/docs-ui/facade$': resolve('node_modules/@univerjs/docs-ui/lib/facade.js'),
        // @wendellhu/redi react-bindings (Univer 内部依赖)
        '@wendellhu/redi/react-bindings$': resolve('node_modules/@wendellhu/redi/dist/cjs/react-bindings/index.js'),
        // 纯 ESM 包 (只有 exports 字段，无 main/module，Webpack 4 无法解析)
        'cjk-regex$': resolve('node_modules/cjk-regex/lib/index.js'),
        'regexp-util$': resolve('node_modules/regexp-util/lib/index.js'),
        'unicode-regex$': resolve('node_modules/unicode-regex/lib/index.js')
      }, univerAliases)
    },
    plugins: [
      // http://doc.ruoyi.vip/ruoyi-vue/other/faq.html#使用gzip解压缩静态文件
      new CompressionPlugin({
        cache: false,                                  // 不启用文件缓存
        test: /\.(js|css|html|jpe?g|png|gif|svg)?$/i,  // 压缩文件格式
        filename: '[path][base].gz[query]',            // 压缩后的文件名
        algorithm: 'gzip',                             // 使用gzip压缩
        minRatio: 0.8,                                 // 压缩比例，小于 80% 的文件不会被压缩
        deleteOriginalAssets: false                    // 压缩后删除原文件
      })
    ],
  },
  chainWebpack(config) {
    config.plugins.delete('preload') // TODO: need test
    config.plugins.delete('prefetch') // TODO: need test

    // ===== 修复 .mjs 文件在 Webpack 4 下 strict ESM → "require is not defined" =====
    // Webpack 4 默认把 .mjs 扩展名自动判定为 module.type = "javascript/esm"
    // (strict ESM，禁用 require/module.exports)。但 class-variance-authority、
    // @radix-ui/react-* 等 Univer 依赖的 .mjs 会 import clsx 等 CJS 包，
    // Webpack 在 ESM↔CJS interop 时会生成 require() 桥接 → strict ESM 下崩。
    // 解决：为 node_modules 下的所有 .mjs 加一条"纯类型"规则（不加 loader），
    // 显式声明 module.type = "javascript/auto"（同时允许 ESM import/export 与
    // CJS require/module.exports）。babel 转译仍走 Vue CLI 默认 js rule
    // (test: /\.m?jsx?$/)，不双转译、不影响行为。
    config.module
      .rule('mjs-type-fix')
      .test(/\.mjs$/)
      .include.add(/node_modules[\\/]/).end()
      .type('javascript/auto')
      .end()

    // .mjs 后缀加入 resolve.extensions
    config.resolve.extensions.prepend('.mjs')

    // set svg-sprite-loader
    config.module
      .rule('svg')
      .exclude.add(resolve('src/assets/icons'))
      .end()
    config.module
      .rule('icons')
      .test(/\.svg$/)
      .include.add(resolve('src/assets/icons'))
      .end()
      .use('svg-sprite-loader')
      .loader('svg-sprite-loader')
      .options({
        symbolId: 'icon-[name]'
      })
      .end()

    config.when(process.env.NODE_ENV !== 'development', config => {
          config
            .plugin('ScriptExtHtmlWebpackPlugin')
            .after('html')
            .use('script-ext-html-webpack-plugin', [{
            // `runtime` must same as runtimeChunk name. default is `runtime`
              inline: /runtime\..*\.js$/
            }])
            .end()

          config.optimization.splitChunks({
            chunks: 'all',
            cacheGroups: {
              libs: {
                name: 'chunk-libs',
                test: /[\\/]node_modules[\\/]/,
                priority: 10,
                chunks: 'initial' // only package third parties that are initially dependent
              },
              elementUI: {
                name: 'chunk-elementUI', // split elementUI into a single package
                test: /[\\/]node_modules[\\/]_?element-ui(.*)/, // in order to adapt to cnpm
                priority: 20 // the weight needs to be larger than libs and app or it will be packaged into libs or app
              },
              commons: {
                name: 'chunk-commons',
                test: resolve('src/components'), // can customize your rules
                minChunks: 3, //  minimum common number
                priority: 5,
                reuseExistingChunk: true
              }
            }
          })
          config.optimization.runtimeChunk('single')
    })
  }
}
