
def test_import():
    from get_fve_token import GraphMailClient, fetch_token_simple
    assert callable(fetch_token_simple)
