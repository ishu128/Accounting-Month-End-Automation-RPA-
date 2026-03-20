def is_browser_alive(driver):
    """
    Checks if WebDriver session is still valid.
    Returns False if browser was closed manually.
    """
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False
