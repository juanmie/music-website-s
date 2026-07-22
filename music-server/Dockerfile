# ================================
# Stage 1: Build Spring Boot JAR
# ================================
FROM maven:3.8-openjdk-8 AS builder
WORKDIR /build
COPY pom.xml .
RUN mvn dependency:go-offline -B
COPY src ./src
RUN mvn package -DskipTests -B

# ================================
# Stage 2: Runtime
# ================================
FROM openjdk:8-jre-slim
WORKDIR /app
COPY --from=builder /build/target/*.jar app.jar

# 暴露端口
EXPOSE 8888

# 启动参数：使用 prod 配置，JVM 内存限制
ENTRYPOINT ["java", "-Xms256m", "-Xmx512m", "-Dspring.profiles.active=prod", "-jar", "app.jar"]
