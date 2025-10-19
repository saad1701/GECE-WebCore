run:
	uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
test:
	pytest -q
docker-build:
	docker build -t gece-backend:latest .
docker-run:
	docker run -p 8000:8000 gece-backend:latest