## process-gpt-office-mcp

HWPX 템플릿을 채워 결과 파일을 생성하는 MCP 서버입니다.  
향후 DOCX, PPT 생성 기능을 추가할 수 있도록 설계했습니다.

### 실행 방법
```bash
python main.py
```

기본 포트: `1192`

### 환경 변수
`.env`에 API 키만 설정합니다.
```
OPENAI_API_KEY=
GOOGLE_API_KEY=
```

### Docker
```bash
docker build -t process-gpt-office-mcp .
docker run --rm -p 1192:1192 process-gpt-office-mcp
```

### 배포 (GitHub Actions)
워크플로:
- `.github/workflows/deploy.yaml` (main/master push)
- `.github/workflows/deploy-prod.yaml` (release 생성)

이미지:
`ghcr.io/uengine-oss/process-gpt-office-mcp:<tag>`

### Kubernetes (GitOps)
`process-gpt-k8s` 레포에 아래 리소스를 추가합니다.
- `deployments/process-gpt-office-mcp-deployment.yaml`
- `services/process-gpt-office-mcp-service.yaml`
