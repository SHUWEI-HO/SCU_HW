# FROM UBUNTU
# 冒號可指定版本
FROM ubuntu:22.04                     
RUN apt update && apt install -y nginx

COPY index.html /vaer/ww/html/

# COMMON
EXPOSE 80
ENTRYPOINT ["nginx", "-g", "daemon off;"]
