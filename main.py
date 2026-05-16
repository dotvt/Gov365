from fastapi import FastAPI
from pydantic import BaseModel
import requests

app = FastAPI()

# =========================
# REQUEST MODEL
# =========================

class GraphRequest(BaseModel):
    tenant_id: str
    client_id: str
    client_secret: str

# =========================
# TOKEN FUNCTION
# =========================

def get_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(url, data=payload)
    return response.json()

# =========================
# GRAPH CALL
# =========================

def test_graph_access(access_token):
    url = "https://graph.microsoft.com/v1.0/organization"

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)

    return {
        "status_code": response.status_code,
        "response": response.json() if response.content else {}
    }

# =========================
# MAIN ENDPOINT
# =========================

@app.post("/test-graph")
def test_graph(req: GraphRequest):

    # 1. GET TOKEN
    token_response = get_token(
        req.tenant_id,
        req.client_id,
        req.client_secret
    )

    if "access_token" not in token_response:
        return {
            "status": "FAILED_TOKEN",
            "details": token_response
        }

    access_token = token_response["access_token"]

    # 2. CALL GRAPH
    graph_result = test_graph_access(access_token)

    # 3. RETURN RESULT
    if graph_result["status_code"] == 200:
        return {
            "status": "SUCCESS",
            "message": "Graph connection successful",
            "graph_data": graph_result["response"]
        }

    return {
        "status": "FAILED_GRAPH",
        "message": "Graph API call failed",
        "graph_data": graph_result
    }
