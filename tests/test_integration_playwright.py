def test_login_y_buscar_ot_mock(mocker):
    mock_context = mocker.Mock()
    mock_page = mocker.Mock()
    mock_context.new_page.return_value = mock_page
    mock_page.locator().first.count.return_value = 1
    mock_page.locator().first.inner_text.return_value = "VPE 12345"
    mock_page.locator().nth.return_value.get_attribute.return_value = "/instrumentoDetalle.do?id=1"
    from src.portal.scraper import login_y_buscar_ot
    _, meta, hrefs = login_y_buscar_ot(mock_context, "u", "p", "307-12345")
    assert meta["vpe"] == "12345"
    assert hrefs
