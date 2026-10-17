from pathlib import Path

from dotenv import load_dotenv

# Settings such as TDS_S3_BUCKET and the AWS keys; real environment variables take precedence.
for _env in (Path(__file__).resolve().parents[1] / ".env", Path(__file__).resolve().parent / ".env"):
    load_dotenv(_env)

import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api.comparison import router as comparison_router

app = FastAPI(
    title="TDS Post-Filing Change Comparison Tool",
    description="Deterministic accounting comparison engine for TDS post-filing audit",
    version="1.0.0",
)

# Sites allowed to call the API from a browser: FRONTEND_ORIGIN, comma-separated (e.g. https://tds.vercel.app),
# when the frontend is hosted on another domain; any site when unset (local development)
_origins = [o.strip().rstrip("/") for o in os.environ.get("FRONTEND_ORIGIN", "").split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins or ["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(comparison_router)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
