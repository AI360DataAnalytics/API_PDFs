from fastapi import FastAPI
import uvicorn

from api.v1.endpoints.document import router as document_router

app = FastAPI()

app.include_router(document_router, prefix="/api/v1")


if __name__ == "__main__":
    
    uvicorn.run(app, host="0.0.0.0", port=8000)