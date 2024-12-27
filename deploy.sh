#!/bin/bash

# 设置 Docker Hub 用户名
DOCKER_USERNAME=codershaochong
IMAGE_NAME=excel-compare
VERSION=latest

# 停止并删除旧容器（如果存在）
docker stop excel-compare || true
docker rm excel-compare || true

# 构建新镜像
docker build --platform=linux/amd64 -t ${IMAGE_NAME} .

# 给镜像打标签
docker tag ${IMAGE_NAME} ${DOCKER_USERNAME}/${IMAGE_NAME}:${VERSION}

# 运行新容器
docker run -d -p 80:80 --name excel-compare ${DOCKER_USERNAME}/${IMAGE_NAME}:${VERSION}

# 推送到 Docker Hub
docker push ${DOCKER_USERNAME}/${IMAGE_NAME}:${VERSION} 