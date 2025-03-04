from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def home():
    return {"message": "Deployment works!"}

@app.get("/health")
def health_check():
    return {"status": "ok"}
