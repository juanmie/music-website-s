#!/bin/bash
# ================================================
# 服务器环境初始化脚本 - 手工部署方案
# 服务器: CentOS 8.5 64位 | 2核2G
# 用途: 安装 JDK 8、MySQL、Redis、MinIO、Nginx、Node
# ================================================
set -e

echo "=========================================="
echo "  音乐网站 - 服务器环境初始化"
echo "  CentOS 8.5 | 2核2G | 手工部署"
echo "=========================================="

# ============================
# 1. 基础依赖
# ============================
echo ""
echo "========== 1/7 安装基础依赖 =========="
dnf install -y epel-release
dnf install -y wget curl tar gcc make

# ============================
# 2. JDK 8
# ============================
echo ""
echo "========== 2/7 安装 JDK 8 =========="
if java -version 2>&1 | grep -q "1.8"; then
    echo "JDK 8 已安装，跳过"
else
    dnf install -y java-1.8.0-openjdk java-1.8.0-openjdk-devel
    echo "JDK 8 安装完成"
fi
java -version 2>&1 | head -1

# ============================
# 3. MySQL 8.0
# ============================
echo ""
echo "========== 3/7 安装 MySQL 8.0 =========="
if command -v mysqld &>/dev/null; then
    echo "MySQL 已安装，跳过"
else
    # 使用 MySQL 官方仓库
    dnf install -y https://dev.mysql.com/get/mysql80-community-release-el8-9.noarch.rpm
    dnf install -y mysql-community-server --nogpgcheck
    
    # 启动 MySQL
    systemctl start mysqld
    systemctl enable mysqld
    
    # 获取临时密码
    TEMP_PWD=$(grep 'temporary password' /var/log/mysqld.log | tail -1 | awk '{print $NF}')
    echo ""
    echo "=============================="
    echo "  MySQL 临时密码: $TEMP_PWD"
    echo "  请立即执行: mysql -u root -p'$TEMP_PWD'"
    echo "  然后修改密码: ALTER USER 'root'@'localhost' IDENTIFIED BY '你的新密码';"
    echo "=============================="
    echo ""
fi

# ============================
# 4. Redis 5
# ============================
echo ""
echo "========== 4/7 安装 Redis =========="
if command -v redis-server &>/dev/null; then
    echo "Redis 已安装，跳过"
else
    dnf install -y redis
    systemctl start redis
    systemctl enable redis
fi
redis-server --version

# ============================
# 5. MinIO
# ============================
echo ""
echo "========== 5/7 安装 MinIO =========="
if command -v minio &>/dev/null; then
    echo "MinIO 已安装，跳过"
else
    wget https://dl.min.io/server/minio/release/linux-amd64/minio -O /usr/local/bin/minio
    chmod +x /usr/local/bin/minio
    
    # 创建 MinIO 数据目录
    mkdir -p /data/minio
    
    # 创建 MinIO 系统用户
    useradd -r -s /sbin/nologin minio-user 2>/dev/null || true
    chown minio-user:minio-user /data/minio
fi
echo "MinIO 安装完成"

# ============================
# 6. Nginx
# ============================
echo ""
echo "========== 6/7 安装 Nginx =========="
if command -v nginx &>/dev/null; then
    echo "Nginx 已安装，跳过"
else
    dnf install -y nginx
    systemctl enable nginx
fi
nginx -v

# ============================
# 7. Node.js 14 (用于构建前端)
# ============================
echo ""
echo "========== 7/7 安装 Node.js 14 =========="
if command -v node &>/dev/null && node -v | grep -q "v14"; then
    echo "Node.js 14 已安装，跳过"
else
    curl -fsSL https://rpm.nodesource.com/setup_14.x | bash -
    dnf install -y nodejs
fi
node -v
npm -v

# ============================
# 8. 创建应用目录
# ============================
echo ""
echo "========== 创建应用目录 =========="
mkdir -p /opt/music-website
mkdir -p /opt/music-website/app
mkdir -p /opt/music-website/logs

echo ""
echo "=========================================="
echo "  环境初始化完成！"
echo "=========================================="
echo "  JDK:      $(java -version 2>&1 | head -1)"
echo "  MySQL:    $(mysqld --version 2>&1 | head -1)"
echo "  Redis:    $(redis-server --version)"
echo "  Nginx:    $(nginx -v 2>&1)"
echo "  Node.js:  $(node -v)"
echo "  MinIO:    /usr/local/bin/minio"
echo ""
echo "  ⚠️  重要：请先完成 MySQL 密码修改！"
echo "=========================================="
