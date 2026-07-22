#!/bin/bash
# ================================================
# 腾讯云服务器初始化脚本
# 服务器: CentOS 8.5 64位 | IP: 49.235.120.3
# 用途: 安装 Docker + Docker Compose，初始化项目
# ================================================

set -e

echo "========== 1. 安装 Docker =========="
# 卸载旧版本（如有）
yum remove -y docker docker-client docker-client-latest docker-common \
  docker-latest docker-latest-logrotate docker-logrotate docker-engine \
  podman runc 2>/dev/null || true

# 安装依赖
yum install -y yum-utils device-mapper-persistent-data lvm2

# 添加 Docker 仓库（使用阿里云镜像加速）
yum-config-manager --add-repo https://mirrors.aliyun.com/docker-ce/linux/centos/docker-ce.repo
sed -i 's+download.docker.com+mirrors.aliyun.com/docker-ce+' /etc/yum.repos.d/docker-ce.repo

# 安装 Docker
yum install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin

# 启动 Docker 并设置开机自启
systemctl start docker
systemctl enable docker

echo "========== 2. 配置 Docker 镜像加速 =========="
mkdir -p /etc/docker
cat > /etc/docker/daemon.json << 'EOF'
{
  "registry-mirrors": [
    "https://mirror.ccs.tencentyun.com",
    "https://docker.mirrors.ustc.edu.cn"
  ]
}
EOF
systemctl daemon-reload
systemctl restart docker

echo "========== 3. 拉取项目代码 =========="
mkdir -p /opt/music-website
cd /opt/music-website

# 如果是首次运行，需要克隆代码
if [ ! -f "docker-compose.yml" ]; then
  echo "请将代码推送到 GitHub 后，手动克隆到 /opt/music-website"
  echo "git clone https://github.com/YOUR_USERNAME/music-website-s.git /opt/music-website"
  exit 1
fi

echo "========== 4. 构建并启动所有服务 =========="
docker compose up -d --build

echo "========== 5. 查看服务状态 =========="
docker compose ps

echo ""
echo "============================================"
echo "  部署完成！"
echo "============================================"
echo "  用户端前端:  http://49.235.120.3:8080"
echo "  管理后台:    http://49.235.120.3:8081"
echo "  后端 API:    http://49.235.120.3:8888"
echo "  MinIO 控制台: http://49.235.120.3:9000"
echo "============================================"
