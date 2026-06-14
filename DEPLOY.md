# 실서버 배포 가이드 (VPS + 자동 HTTPS)

VPS(AWS EC2 · 네이버클라우드 · Vultr 등)에 **Docker Compose**로 올리는 방법.
앞단의 **Caddy**가 Let's Encrypt 인증서를 자동 발급/갱신해 HTTPS를 붙여준다.
**도메인이 아직 없어도** `sslip.io` 트릭으로 진짜 인증서를 받을 수 있다.

---

## 0. 준비물

- 공인 IP가 있는 리눅스 VPS (Ubuntu 22.04+ 권장)
- 방화벽/보안그룹에서 **80, 443 포트 인바운드 허용** (Let's Encrypt 인증에 80 필요)
- (선택) 챗봇용 `ANTHROPIC_API_KEY`

> ⚠️ **개인정보 주의**: 간호사 명단·근무가 저장된다. 외부 클라우드 VPS에 두는 것이
> 병원 보안정책상 허용되는지 먼저 확인할 것. 가능하면 접근을 사내망/지정 IP로 제한.

---

## 1. 코드 올리기

먼저 로컬에서 이번 배포 파일들을 깃에 커밋·푸시한다(아래 "변경 커밋" 참고).
그다음 VPS에서:

```bash
git clone https://github.com/CheonSeokHee/nurse-duty-scheduler.git
cd nurse-duty-scheduler
```

## 2. Docker 설치 (Ubuntu)

```bash
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker $USER   # 로그아웃 후 재로그인하면 sudo 없이 docker 사용
```

## 3. 접속 주소(SITE_ADDRESS) 정하기

### 도메인이 없을 때 — sslip.io (무료, 즉시)
서버 공인 IP가 `203.0.113.45` 라면 → 점을 하이픈으로 바꿔서:

```
203-0-113-45.sslip.io
```

이 주소가 자동으로 그 IP를 가리키고, Caddy가 여기에 대해 정식 HTTPS 인증서를 받는다.

### 도메인이 생기면
DNS A 레코드를 서버 IP로 연결한 뒤 `SITE_ADDRESS=duty.example.com` 으로 바꾸면 끝.

## 4. (선택) 챗봇 키

```bash
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env
```

## 5. 실행

```bash
SITE_ADDRESS=203-0-113-45.sslip.io docker compose up -d --build
```

- 컨테이너 2개(app, caddy)가 뜬다. 인증서 발급에 수십 초 걸릴 수 있다.
- 브라우저에서 **https://203-0-113-45.sslip.io** 접속 → **첫 관리자 계정 만들기** 화면.

상태 확인:
```bash
docker compose ps
docker compose logs -f caddy   # 인증서 발급 로그
docker compose logs -f app
```

---

## 운영

**업데이트(코드 변경 반영):**
```bash
git pull
SITE_ADDRESS=203-0-113-45.sslip.io docker compose up -d --build
```

**백업** — 모든 데이터(명단·이력·계정)는 `dutydata` 볼륨의 `db.json` 하나에 있다:
```bash
docker compose cp app:/app/server/data/db.json ./db-backup-$(date +%F).json
```

**중지 / 재시작:**
```bash
docker compose down       # 중지(볼륨=데이터는 유지)
docker compose restart
```

---

## 보안 체크리스트 (배포 후)

- [x] HTTPS (Caddy 자동) · [x] Secure/HttpOnly 쿠키 · [x] 비번 scrypt 해시 · [x] 로그인 시도 제한
- [ ] **접근 제한**: 가능하면 보안그룹에서 80/443을 사내/지정 IP로만 허용 (개인정보 보호)
- [ ] **관리자 비번**을 강하게, 첫 로그인 후 불필요한 테스트 계정 정리
- [ ] 정기 **백업**(위 db.json) — 가능하면 자동화(cron)
- [ ] 서버 OS·Docker 정기 업데이트

## 규모가 커지면

- 파일 저장소(`db.json`) → **PostgreSQL** 로 교체 (`server/store.js`만 바꾸면 됨)
- 병동별 접근 권한, 감사 로그, 다중 인스턴스 등
