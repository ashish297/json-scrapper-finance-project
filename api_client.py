import requests


def fetch_officers_by_perm_id(perm_id: str, cookies: dict, headers: dict, timeout_seconds: int = 30) -> requests.Response:
    """Call the OfficersDirectors API for a given Organization PermID and return the raw response.

    Keeping this isolated makes it easy to swap/modify the request without touching the rest of the pipeline.
    """
    params = {
        'fields': 'full',
        'isPublic': 'true',
        'lang': 'en-US',
        'oapermid': str(perm_id),
    }
    return requests.get(
        'https://workspace.refinitiv.com/Apps/OfficersDirectors/officers',
        params=params,
        cookies=cookies,
        headers=headers,
        timeout=timeout_seconds,
    )


