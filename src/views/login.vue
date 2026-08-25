<template>
  <div class="login">
    <!-- 背景层：粒子 + 光晕 + 网格 -->
    <div class="login-bg">
      <canvas ref="particleCanvas" class="particle-canvas"></canvas>
      <div class="bg-glow bg-glow--1"></div>
      <div class="bg-glow bg-glow--2"></div>
      <div class="grid-overlay"></div>
    </div>

    <!-- 登录卡片（毛玻璃） -->
    <el-form ref="loginForm" :model="loginForm" :rules="loginRules" class="login-form">
      <div class="login-form__header">
        <div class="logo-mark">
          <span class="logo-mark__core"></span>
        </div>
        <h3 class="title">{{title}}</h3>
        <p class="subtitle">数据驱动未来</p>
      </div>

      <el-form-item prop="username">
        <el-input
          v-model="loginForm.username"
          type="text"
          auto-complete="off"
          placeholder="账号"
        >
          <svg-icon slot="prefix" icon-class="user" class="el-input__icon input-icon" />
        </el-input>
      </el-form-item>
      <el-form-item prop="password">
        <el-input
          v-model="loginForm.password"
          type="password"
          auto-complete="off"
          placeholder="密码"
          @keyup.enter.native="handleLogin"
        >
          <svg-icon slot="prefix" icon-class="password" class="el-input__icon input-icon" />
        </el-input>
      </el-form-item>
      <el-form-item prop="code" v-if="captchaEnabled">
        <el-input
          v-model="loginForm.code"
          auto-complete="off"
          placeholder="验证码"
          style="width: 63%"
          @keyup.enter.native="handleLogin"
        >
          <svg-icon slot="prefix" icon-class="validCode" class="el-input__icon input-icon" />
        </el-input>
        <div class="login-code">
          <img :src="codeUrl" @click="getCode" class="login-code-img"/>
        </div>
      </el-form-item>
      <el-checkbox v-model="loginForm.rememberMe" class="remember-me">记住密码</el-checkbox>
      <el-form-item style="width:100%;">
        <el-button
          :loading="loading"
          size="medium"
          type="primary"
          class="login-btn"
          @click.native.prevent="handleLogin"
        >
          <span v-if="!loading">登 录</span>
          <span v-else>登 录 中...</span>
        </el-button>
        <div style="float: right;" v-if="register">
          <router-link class="link-type" :to="'/register'">立即注册</router-link>
        </div>
      </el-form-item>
    </el-form>

    <!--  底部  -->
    <div class="el-login-footer">
      <span>{{ footerContent }}</span>
    </div>
  </div>
</template>

<script>
import { getCodeImg } from "@/api/login"
import Cookies from "js-cookie"
import { encrypt, decrypt } from '@/utils/jsencrypt'
import defaultSettings from '@/settings'

export default {
  name: "Login",
  data() {
    return {
      title: process.env.VUE_APP_TITLE,
      footerContent: defaultSettings.footerContent,
      codeUrl: "",
      loginForm: {
        username: "admin",
        password: "admin123",
        rememberMe: false,
        code: "",
        uuid: ""
      },
      loginRules: {
        username: [
          { required: true, trigger: "blur", message: "请输入您的账号" }
        ],
        password: [
          { required: true, trigger: "blur", message: "请输入您的密码" }
        ],
        code: [{ required: true, trigger: "change", message: "请输入验证码" }]
      },
      loading: false,
      // 验证码开关
      captchaEnabled: true,
      // 注册开关
      register: false,
      redirect: undefined,
      // 粒子动画
      particles: [],
      particleCtx: null,
      particleWidth: 0,
      particleHeight: 0,
      particleTimer: null
    }
  },
  watch: {
    $route: {
      handler: function(route) {
        this.redirect = route.query && route.query.redirect
      },
      immediate: true
    }
  },
  created() {
    this.getCode()
    this.getCookie()
  },
  mounted() {
    this.initParticles()
    window.addEventListener('resize', this.handleResize)
  },
  beforeDestroy() {
    window.removeEventListener('resize', this.handleResize)
    if (this.particleTimer) cancelAnimationFrame(this.particleTimer)
  },
  methods: {
    getCode() {
      getCodeImg().then(res => {
        this.captchaEnabled = res.captchaEnabled === undefined ? true : res.captchaEnabled
        if (this.captchaEnabled) {
          this.codeUrl = "data:image/gif;base64," + res.img
          this.loginForm.uuid = res.uuid
        }
      })
    },
    getCookie() {
      const username = Cookies.get("username")
      const password = Cookies.get("password")
      const rememberMe = Cookies.get('rememberMe')
      this.loginForm = {
        username: username === undefined ? this.loginForm.username : username,
        password: password === undefined ? this.loginForm.password : decrypt(password),
        rememberMe: rememberMe === undefined ? false : Boolean(rememberMe)
      }
    },
    handleLogin() {
      this.$refs.loginForm.validate(valid => {
        if (valid) {
          this.loading = true
          if (this.loginForm.rememberMe) {
            Cookies.set("username", this.loginForm.username, { expires: 30 })
            Cookies.set("password", encrypt(this.loginForm.password), { expires: 30 })
            Cookies.set('rememberMe', this.loginForm.rememberMe, { expires: 30 })
          } else {
            Cookies.remove("username")
            Cookies.remove("password")
            Cookies.remove('rememberMe')
          }
          this.$store.dispatch("Login", this.loginForm).then(() => {
            this.$router.push({ path: this.redirect || "/" }).catch(()=>{})
          }).catch(() => {
            this.loading = false
            if (this.captchaEnabled) {
              this.getCode()
            }
          })
        }
      })
    },
    // ===== 粒子动画 =====
    initParticles() {
      const canvas = this.$refs.particleCanvas
      if (!canvas) return
      const ctx = canvas.getContext('2d')
      this.particleCtx = ctx
      this.handleResize()

      const count = Math.min(80, Math.floor(this.particleWidth / 18))
      this.particles = Array.from({ length: count }, () => this.createParticle())

      const step = () => {
        this.drawParticles()
        this.particleTimer = requestAnimationFrame(step)
      }
      this.particleTimer = requestAnimationFrame(step)
    },
    handleResize() {
      const canvas = this.$refs.particleCanvas
      if (!canvas) return
      const rect = canvas.parentElement.getBoundingClientRect()
      const dpr = window.devicePixelRatio || 1
      canvas.width = rect.width * dpr
      canvas.height = rect.height * dpr
      canvas.style.width = rect.width + 'px'
      canvas.style.height = rect.height + 'px'
      this.particleWidth = canvas.width
      this.particleHeight = canvas.height
      this.particleCtx = canvas.getContext('2d')
      this.particleCtx.scale(dpr, dpr)
    },
    createParticle() {
      return {
        x: Math.random() * this.particleWidth,
        y: Math.random() * this.particleHeight,
        vx: (Math.random() - 0.5) * 0.4,
        vy: (Math.random() - 0.5) * 0.4,
        r: Math.random() * 2 + 0.5,
        alpha: Math.random() * 0.5 + 0.2
      }
    },
    drawParticles() {
      const ctx = this.particleCtx
      if (!ctx) return
      const w = this.particleWidth
      const h = this.particleHeight
      ctx.clearRect(0, 0, w, h)

      // 连线
      for (let i = 0; i < this.particles.length; i++) {
        const p = this.particles[i]
        p.x += p.vx
        p.y += p.vy
        if (p.x < 0 || p.x > w) p.vx *= -1
        if (p.y < 0 || p.y > h) p.vy *= -1

        for (let j = i + 1; j < this.particles.length; j++) {
          const q = this.particles[j]
          const dx = p.x - q.x
          const dy = p.y - q.y
          const dist = Math.sqrt(dx * dx + dy * dy)
          if (dist < 120) {
            ctx.strokeStyle = 'rgba(64, 180, 255, ' + (0.14 * (1 - dist / 120)) + ')'
            ctx.lineWidth = 0.6
            ctx.beginPath()
            ctx.moveTo(p.x, p.y)
            ctx.lineTo(q.x, q.y)
            ctx.stroke()
          }
        }
      }

      // 绘制粒子
      this.particles.forEach(p => {
        ctx.fillStyle = 'rgba(120, 210, 255, ' + p.alpha + ')'
        ctx.beginPath()
        ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2)
        ctx.fill()
      })
    }
  }
}
</script>

<style rel="stylesheet/scss" lang="scss" scoped>
.login {
  position: relative;
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100%;
  overflow: hidden;
  background: radial-gradient(ellipse at 20% 20%, #0a1a3a 0%, #060d1f 45%, #03060f 100%);
}

/* ===== 背景层 ===== */
.login-bg {
  position: absolute;
  inset: 0;
  overflow: hidden;
}
.particle-canvas {
  position: absolute;
  inset: 0;
  z-index: 1;
}
.bg-glow {
  position: absolute;
  border-radius: 50%;
  filter: blur(90px);
  opacity: 0.55;
  z-index: 0;
  animation: float 12s ease-in-out infinite;
}
.bg-glow--1 {
  width: 420px;
  height: 420px;
  left: -80px;
  top: -60px;
  background: radial-gradient(circle, rgba(37, 99, 235, 0.7), transparent 70%);
}
.bg-glow--2 {
  width: 520px;
  height: 520px;
  right: -120px;
  bottom: -120px;
  background: radial-gradient(circle, rgba(6, 182, 212, 0.55), transparent 70%);
  animation-delay: -6s;
}
@keyframes float {
  0%, 100% { transform: translate(0, 0) scale(1); }
  50% { transform: translate(30px, -30px) scale(1.08); }
}
.grid-overlay {
  position: absolute;
  inset: 0;
  z-index: 0;
  background-image:
    linear-gradient(rgba(64, 180, 255, 0.06) 1px, transparent 1px),
    linear-gradient(90deg, rgba(64, 180, 255, 0.06) 1px, transparent 1px);
  background-size: 46px 46px;
  mask-image: radial-gradient(ellipse at center, rgba(0,0,0,0.9), transparent 75%);
  -webkit-mask-image: radial-gradient(ellipse at center, rgba(0,0,0,0.9), transparent 75%);
}

/* ===== 登录卡片 ===== */
.title {
  margin: 0;
  text-align: center;
  color: #e8f1ff;
  font-size: 26px;
  font-weight: 700;
  letter-spacing: 3px;
  background: linear-gradient(120deg, #7dd3fc, #38bdf8, #818cf8);
  -webkit-background-clip: text;
  background-clip: text;
  -webkit-text-fill-color: transparent;
}
.subtitle {
  margin: 10px 0 26px;
  text-align: center;
  color: rgba(148, 178, 214, 0.7);
  font-size: 13px;
  letter-spacing: 2px;
}
.logo-mark {
  position: relative;
  width: 56px;
  height: 56px;
  margin: 0 auto 18px;
  border-radius: 16px;
  background: linear-gradient(135deg, rgba(56, 189, 248, 0.25), rgba(129, 140, 248, 0.25));
  border: 1px solid rgba(120, 210, 255, 0.35);
  box-shadow: 0 0 24px rgba(56, 189, 248, 0.45);
  display: flex;
  align-items: center;
  justify-content: center;
  animation: pulse 2.6s ease-in-out infinite;
}
.logo-mark__core {
  width: 22px;
  height: 22px;
  border-radius: 50%;
  background: radial-gradient(circle at 35% 35%, #bae6fd, #38bdf8 55%, #6366f1);
  box-shadow: 0 0 18px rgba(56, 189, 248, 0.9);
}
@keyframes pulse {
  0%, 100% { box-shadow: 0 0 24px rgba(56, 189, 248, 0.45); }
  50% { box-shadow: 0 0 40px rgba(56, 189, 248, 0.75); }
}

.login-form {
  position: relative;
  z-index: 2;
  border-radius: 16px;
  background: rgba(13, 25, 48, 0.6);
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
  border: 1px solid rgba(120, 180, 255, 0.22);
  box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5), 0 0 0 1px rgba(56, 189, 248, 0.08) inset;
  width: 400px;
  padding: 36px 32px 16px 32px;

  .el-input {
    height: 44px;
    input {
      height: 44px;
      background: rgba(20, 34, 62, 0.6);
      border: 1px solid rgba(120, 180, 255, 0.2);
      border-radius: 8px;
      color: #e8f1ff;
      padding-left: 38px;
      transition: border-color 0.25s, box-shadow 0.25s;
      &::placeholder {
        color: rgba(148, 178, 214, 0.5);
      }
      &:focus {
        border-color: rgba(56, 189, 248, 0.7);
        box-shadow: 0 0 0 3px rgba(56, 189, 248, 0.18);
      }
    }
  }
  .input-icon {
    height: 44px;
    width: 14px;
    margin-left: 2px;
    color: #7dd3fc;
  }
}
.remember-me {
  margin: 0 0 25px 0;
  color: rgba(180, 200, 230, 0.75);
}
.login-btn {
  width: 100%;
  height: 44px;
  border: none;
  border-radius: 8px;
  font-size: 16px;
  letter-spacing: 6px;
  background: linear-gradient(120deg, #2563eb, #38bdf8);
  box-shadow: 0 8px 24px rgba(37, 99, 235, 0.45);
  transition: transform 0.15s, box-shadow 0.25s;
  &:hover {
    transform: translateY(-2px);
    box-shadow: 0 12px 32px rgba(56, 189, 248, 0.55);
  }
}
.link-type {
  color: #7dd3fc;
  font-size: 13px;
  &:hover {
    color: #38bdf8;
  }
}
.login-tip {
  font-size: 13px;
  text-align: center;
  color: #bfbfbf;
}
.login-code {
  width: 33%;
  height: 44px;
  float: right;
  img {
    cursor: pointer;
    vertical-align: middle;
    height: 44px;
    border-radius: 8px;
  }
}
.login-code-img {
  height: 44px;
}
.el-login-footer {
  position: fixed;
  bottom: 0;
  width: 100%;
  height: 40px;
  line-height: 40px;
  text-align: center;
  color: rgba(160, 185, 215, 0.6);
  font-family: Arial;
  font-size: 12px;
  letter-spacing: 1px;
  z-index: 2;
}
</style>
