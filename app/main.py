from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api.comparison import router as comparison_router

app = FastAPI(
    title="TDS Post-Filing Change Comparison Tool",
    description="Deterministic accounting comparison engine for TDS post-filing audit",
    version="1.0.0",
)

# Enable CORS for frontend integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(comparison_router)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
