// Test-only isolated entry. Uses the production page, Vue and Univer.
import Vue from 'vue'
import ElementUI from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css'
import Report from '../../src/views/application/agentReport/index.vue'
import { executeActions } from '../../src/views/application/agentReport/univer-action-executor'
Vue.use(ElementUI)
Vue.prototype.$modal = { msgError: console.error, msgWarning: console.warn, msgSuccess: console.log, confirm: message => Promise.resolve(message) }
Vue.prototype.$message = ElementUI.Message
const app = new Vue({ render: h => h(Report) }).$mount('#app')
window.a1Test = { page: () => app.$children[0], executeActions }
