#!/bin/bash
# ================================================
# 应用部署脚本 - 手工部署方案
# 用途: 构建前端 + 后端，部署到服务器
# 首次部署和更新部署都使用此脚本
# ================================================
set -e

APP_DIR=/opt/music-website
SRC_DIR=/opt/music-website/source
LOG_DIR=/opt/music-website/logs

echo "=========================================="
echo "  音乐网站 - 应用部署"
echo "=========================================="

# ============================
# 1. 拉取最新代码
# ============================
echo ""
echo "========== 1/5 拉取最新代码 =========="
mkdir -p $SRC_DIR
if [ -d "$SRC_DIR/.git" ]; then
    cd $SRC_DIR
    git pull origin main
else
    echo "首次部署：请手动克隆代码到 $SRC_DIR"
    echo "  git clone https://github.com/juanmie/music-website-s.git $SRC_DIR"
    exit 1
fi

# ============================
# 2. 构建后端 JAR
# ============================
echo ""
echo "========== 2/5 构建后端 =========="
cd $SRC_DIR/music-server
./mvnw package -DskipTests -B
cp target/*.jar $APP_DIR/app/music-server.jar
echo "后端构建完成"

# ============================
# 3. 构建前端
# ============================
echo ""
echo "========== 3/5 构建前端 =========="

# 用户端
echo "  构建 music-client..."
cd $SRC_DIR/music-client
npm install --registry=https://registry.npmmirror.com
npm run build
rm -rf $APP_DIR/app/music-client
cp -r dist $APP_DIR/app/music-client

# 管理端
echo "  构建 music-manage..."
cd $SRC_DIR/music-manage
npm install --registry=https://registry.npmmirror.com
npm run build
rm -rf $APP_DIR/app/music-manage
cp -r dist $APP_DIR/app/music-manage

# 需求评估端
echo "  构建 demand-client..."
cd $SRC_DIR/demand-client
npm install --registry=https://registry.npmmirror.com
npm run build
rm -rf $APP_DIR/app/demand-client
cp -r dist $APP_DIR/app/demand-client

echo "前端构建完成"

# ============================
# 4. 配置 Nginx
# ============================
echo ""
echo "========== 4/5 配置 Nginx =========="
cp $SRC_DIR/deploy/nginx/music-client.conf /etc/nginx/conf.d/
cp $SRC_DIR/deploy/nginx/music-manage.conf /etc/nginx/conf.d/
cp $SRC_DIR/deploy/nginx/demand-client.conf /etc/nginx/conf.d/

# 删除 Nginx 默认配置（避免冲突）
rm -f /etc/nginx/conf.d/default.conf

# 测试配置
nginx -t
systemctl reload nginx
echo "Nginx 配置完成"

# ============================
# 5. 配置并启动服务
# ============================
echo ""
echo "========== 5/5 启动服务 =========="

# 安装 systemd 服务文件
cp $SRC_DIR/deploy/systemd/music-server.service /etc/systemd/system/
cp $SRC_DIR/deploy/systemd/minio.service /etc/systemd/system/

# 重新加载 systemd
systemctl daemon-reload

# 启动 MinIO
systemctl start minio
systemctl enable minio
echo "  MinIO 已启动 (端口 9000/9001)"

# 启动后端
systemctl start music-server
systemctl enable music-server
echo "  后端已启动 (端口 8888)"

# 等待后端启动完成
echo "  等待后端启动..."
sleep 15

echo ""
echo "=========================================="
echo "  部署完成！"
echo "=========================================="
echo ""
echo "  用户端前端:   http://49.235.120.3:8080"
echo "  管理后台:     http://49.235.120.3:8081"
echo "  需求评估:     http://49.235.120.3:8082"
echo "  后端 API:     http://49.235.120.3:8888"
echo "  MinIO 控制台:  http://49.235.120.3:9001"
echo ""
echo "  服务管理命令:"
echo "    systemctl status music-server   # 查看后端状态"
echo "    systemctl restart music-server  # 重启后端"
echo "    systemctl status minio          # 查看 MinIO 状态"
echo "    systemctl reload nginx          # 重载 Nginx"
echo "    tail -f $LOG_DIR/server.log     # 查看后端日志"
echo "=========================================="
