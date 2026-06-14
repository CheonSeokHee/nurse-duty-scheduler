# 간호사 듀티표 생성기 — 운영 이미지
FROM node:20-alpine

WORKDIR /app

# 의존성 먼저(캐시 최적화)
COPY package*.json ./
RUN npm install --omit=dev && npm cache clean --force

# 앱 소스
COPY . .

ENV NODE_ENV=production
ENV PORT=3000
EXPOSE 3000

# 데이터(db.json)는 볼륨으로 마운트 → /app/server/data
CMD ["node", "server/index.js"]
