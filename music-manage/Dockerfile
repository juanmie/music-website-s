# ================================
# Stage 1: Build Vue 前端
# ================================
FROM node:14-alpine AS builder
WORKDIR /app
COPY package*.json ./
RUN npm install --registry=https://registry.npmmirror.com
COPY . .
# 覆盖 NODE_HOST 为空字符串，让 API 请求走相对路径（由 Nginx 代理）
ARG NODE_HOST='""'
ENV NODE_HOST=${NODE_HOST}
RUN npm run build

# ================================
# Stage 2: Nginx 运行
# ================================
FROM nginx:1.24-alpine
COPY --from=builder /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
