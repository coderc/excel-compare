# 使用 nginx 作为基础镜像
FROM --platform=linux/amd64 nginx:alpine

# 复制项目文件到 nginx 的默认静态文件目录
COPY index.html /usr/share/nginx/html/
COPY script.js /usr/share/nginx/html/

# 暴露 80 端口
EXPOSE 80

# nginx 会自动启动，所以不需要额外的 CMD 命令 