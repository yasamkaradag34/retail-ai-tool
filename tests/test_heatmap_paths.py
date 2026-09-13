from fastapi.testclient import TestClient

import main


client = TestClient(main.app)


def test_heatmap_page_catalog_filters_by_page_type():
    response = client.get("/api/heatmap/pages", params={"page_type": "plp"})

    assert response.status_code == 200
    data = response.json()
    assert data["total"] == 3
    assert all(page["type"] == "plp" for page in data["pages"])
    assert "/kategori/dizel-enjektorler" in {page["path"] for page in data["pages"]}


def test_exact_page_path_refreshes_heatmap_scope_and_metrics():
    aggregate = client.get("/api/heatmap/data", params={"page": "pdp"}).json()
    exact = client.get(
        "/api/heatmap/data",
        params={"page": "pdp", "path": "/urun/bosch-common-rail-0445110"},
    ).json()

    assert exact["page_context"] == {
        "page_type": "pdp",
        "page_type_label": "Product detail pages",
        "scope_title": "Bosch Common Rail Injector",
        "path": "/urun/bosch-common-rail-0445110",
        "scope": "exact_path",
        "matching_pages": 1,
        "known_path": True,
    }
    assert exact["summary"]["total_sessions"] < aggregate["summary"]["total_sessions"]
    assert exact["summary"]["total_clicks"] < aggregate["summary"]["total_clicks"]
    assert exact["top_elements"][0]["clicks"] < aggregate["top_elements"][0]["clicks"]


def test_custom_full_url_is_normalised_and_page_type_is_inferred():
    response = client.get(
        "/api/heatmap/data",
        params={
            "page": "custom",
            "path": "https://injectormarketing.com/category/turbo-parts?campaign=summer#products",
        },
    )

    assert response.status_code == 200
    context = response.json()["page_context"]
    assert context["page_type"] == "plp"
    assert context["path"] == "/category/turbo-parts"
    assert context["scope"] == "exact_path"
    assert context["known_path"] is False


def test_heatmap_rejects_unsafe_page_path():
    response = client.get("/api/heatmap/data", params={"page": "custom", "path": "/<script>"})

    assert response.status_code == 422
