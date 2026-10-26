"""Reject unavailable Mac interfaces before mutating native objects."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from src.models import DocumentInput, JsonObject
from src.native.javascript import JsObject
from src.presentation import charts, connectors, slideshow, smartart


def test_macos_chart_data_rejection_precedes_navigation_and_workbook_access(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    chart = Mock(spec=JsObject)
    resolve = Mock(return_value=SimpleNamespace(Chart=chart))
    navigate, workbook = Mock(), Mock()
    monkeypatch.setattr(charts, "_get_chart_shape", resolve)
    monkeypatch.setattr(charts, "goto_slide", navigate)
    monkeypatch.setattr(charts, "chart_workbook", workbook)
    params = charts.SetChartDataInput.model_validate(
        {
            "slide_index": 1,
            "shape_name_or_index": "Chart",
            "categories": ["A"],
            "series": [{"name": "Sales", "values": [1.0]}],
        }
    )
    with pytest.raises(NotImplementedError, match="ChartData.Workbook"):
        charts.set_chart_data(params)
    resolve.assert_called_once_with(params, navigate=False)
    navigate.assert_not_called()
    workbook.assert_not_called()
    assert chart.mock_calls == []


def test_macos_chart_read_does_not_activate_embedded_editor(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    chart = Mock(spec=JsObject, ChartData=Mock())
    monkeypatch.setattr(
        charts, "_get_chart_shape", Mock(return_value=Mock(Chart=chart))
    )
    monkeypatch.setattr(
        charts,
        "_read_chart_series",
        Mock(return_value=([""], [{"name": None, "values": [42]}])),
    )
    result = charts.get_chart_data(
        charts.GetChartDataInput(slide_index=1, shape_name_or_index=1)
    )
    assert result["series"] == [{"name": None, "values": [42]}]
    chart.ChartData.Activate.assert_not_called()


def test_macos_chart_creation_does_not_activate_embedded_editor(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    chart = Mock(spec=JsObject, ChartData=Mock())
    session = Mock()
    shapes = session.get_document.return_value.Slides.return_value.Shapes
    shapes.AddChart2.return_value = Mock(Chart=chart)
    monkeypatch.setattr(charts, "get_session", Mock(return_value=session))
    monkeypatch.setattr(charts, "goto_slide", Mock())
    charts.add_chart(charts.AddChartInput(slide_index=1))
    chart.ChartData.Activate.assert_not_called()


@pytest.fixture
def connector_session(monkeypatch: pytest.MonkeyPatch) -> Mock:
    session = Mock()
    monkeypatch.setattr(connectors, "get_session", Mock(return_value=session))
    monkeypatch.setattr(connectors, "goto_slide", Mock())
    return session


@pytest.mark.parametrize("site", ["top", "bottom", "left", "right"])
def test_macos_connector_directions_fail_before_creation(
    connector_session: Mock, monkeypatch: pytest.MonkeyPatch, site: str
) -> None:
    begin = Mock(spec=JsObject, Name="begin", ConnectionSiteCount=4)
    end = Mock(spec=JsObject, Name="end", ConnectionSiteCount=4)
    monkeypatch.setattr(
        connectors, "resolve_shape", Mock(side_effect=[begin, end])
    )
    params = connectors.AddConnectorInput.model_validate(
        {
            "slide_index": 1,
            "begin_shape": "begin",
            "end_shape": "end",
            "begin_site": site,
        }
    )
    with pytest.raises(NotImplementedError, match="numeric connection"):
        connectors.add_connector(params)
    connector_session.get_document().Slides().Shapes.AddConnector.assert_not_called()
    connector_session.get_application.assert_not_called()
    assert begin.mock_calls == end.mock_calls == []


def test_macos_connector_numeric_sites_keep_exact_native_indices(
    connector_session: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    begin = Mock(spec=JsObject, Name="begin", ConnectionSiteCount=4)
    end = Mock(spec=JsObject, Name="end", ConnectionSiteCount=4)
    monkeypatch.setattr(
        connectors, "resolve_shape", Mock(side_effect=[begin, end])
    )
    shapes = connector_session.get_document().Slides().Shapes
    shapes.AddConnector.return_value.Name = "new connector"
    result = connectors.add_connector(
        connectors.AddConnectorInput(
            slide_index=1,
            begin_shape="begin",
            begin_site=3,
            end_shape="end",
            end_site=4,
        )
    )
    assert result["status"] == "success"
    created = shapes.AddConnector.return_value
    created.ConnectorFormat.BeginConnect.assert_called_once_with(begin, 3)
    created.ConnectorFormat.EndConnect.assert_called_once_with(end, 4)
    created.RerouteConnections.assert_called_once_with()


def test_macos_direction_reconnect_fails_before_formatting(
    connector_session: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    connector = Mock(Connector=True)
    target = Mock(spec=JsObject, Name="target", ConnectionSiteCount=4)
    monkeypatch.setattr(
        connectors, "resolve_shape", Mock(side_effect=[connector, target])
    )
    apply = Mock()
    monkeypatch.setattr(connectors, "apply_formatting", apply)
    with pytest.raises(NotImplementedError, match="numeric connection"):
        connectors.format_connector(
            connectors.FormatConnectorInput(
                slide_index=1,
                shape_name_or_index="connector",
                end_shape="target",
                end_site="left",
                weight=3.0,
            )
        )
    apply.assert_not_called()
    connector_session.get_application.assert_not_called()
    assert connector.mock_calls == target.mock_calls == []


def test_non_macos_connector_direction_resolution_is_preserved() -> None:
    sites = Mock(Count=2, side_effect=[(0.0, 50.0), (100.0, 50.0)])
    shape = SimpleNamespace(
        Name="box", Left=0, Top=0, Width=100, Height=100, ConnectionSites=sites
    )
    assert connectors.resolve_site(shape, "right") == 2


@pytest.fixture
def smartart_session(monkeypatch: pytest.MonkeyPatch) -> Mock:
    session = Mock()
    pres = Mock(spec=JsObject, Slides=Mock())
    diagram = Mock(spec=JsObject)
    shape = Mock(Name="Diagram", Type=smartart.MSO_SMART_ART, SmartArt=diagram)
    session.get_document.return_value = pres
    pres.Slides.return_value.Shapes.AddSmartArt.return_value = shape
    session.get_application.return_value.SmartArtLayouts.return_value.Name = (
        "Layout"
    )
    monkeypatch.setattr(smartart, "get_session", Mock(return_value=session))
    monkeypatch.setattr(smartart, "_get_shape", Mock(return_value=shape))
    return session


@pytest.mark.parametrize(
    "options",
    [
        {"node_texts": ["New node"]},
        {"font_name": "Arial"},
        {"font_size": 18.0},
        {"bold": False},
        {"font_color": "#123456"},
    ],
)
def test_macos_smartart_add_node_options_fail_before_shape_creation(
    smartart_session: Mock, options: JsonObject
) -> None:
    params = smartart.AddSmartArtInput.model_validate(
        {"slide_index": 1, **options}
    )
    with pytest.raises(NotImplementedError, match="SmartArt nodes"):
        smartart.add_smartart(params)
    pres = smartart_session.get_document.return_value
    pres.Slides.assert_not_called()
    pres.Slides.return_value.Shapes.AddSmartArt.assert_not_called()
    assert smartart_session.get_application.return_value.mock_calls == []


def test_macos_smartart_add_keeps_diagram_color_and_style(
    smartart_session: Mock,
) -> None:
    result = smartart.add_smartart(
        smartart.AddSmartArtInput(
            slide_index=1, layout_index=2, color_index=3, style_index=4
        )
    )
    app = smartart_session.get_application.return_value
    add = (
        smartart_session.get_document.return_value.Slides().Shapes.AddSmartArt
    )
    add.assert_called_once_with(
        app.SmartArtLayouts.return_value, 50.0, 50.0, 600.0, 400.0
    )
    diagram = add.return_value.SmartArt
    assert diagram.QuickStyle is app.SmartArtQuickStyles.return_value
    assert diagram.Color is app.SmartArtColors.return_value
    assert result["success"] and result["node_count"] is None


@pytest.mark.parametrize(
    "options",
    [
        {"action": "set_text", "node_index": 1, "text": "Changed"},
        {"action": "add_node", "text": "New node"},
        {"action": "delete_node", "node_index": 1},
        {"action": "format_node", "node_index": 1, "bold": False},
        {"action": "format_all_nodes", "font_size": 18.0},
        {"action": "change_layout", "layout_index": 2, "node_index": 1},
    ],
)
def test_macos_smartart_node_actions_fail_without_modification(
    smartart_session: Mock, options: JsonObject
) -> None:
    shapes = smartart_session.get_document.return_value.Slides().Shapes
    shape = shapes.AddSmartArt.return_value
    diagram = shape.SmartArt
    before = dict(vars(diagram))
    params = smartart.ModifySmartArtInput.model_validate(
        {"slide_index": 1, "shape_name_or_index": "Diagram", **options}
    )
    with pytest.raises(NotImplementedError, match="only diagram"):
        smartart.modify_smartart(params)
    assert vars(diagram) == before
    assert diagram.mock_calls == []
    assert smartart_session.get_application.return_value.mock_calls == []


@pytest.mark.parametrize(
    ("action", "field", "collection", "attribute"),
    [
        ("change_layout", "layout_index", "SmartArtLayouts", "Layout"),
        ("change_color", "color_index", "SmartArtColors", "Color"),
        ("change_style", "style_index", "SmartArtQuickStyles", "QuickStyle"),
    ],
)
def test_macos_smartart_diagram_changes_keep_unknown_node_count(
    smartart_session: Mock,
    action: str,
    field: str,
    collection: str,
    attribute: str,
) -> None:
    result = smartart.modify_smartart(
        smartart.ModifySmartArtInput.model_validate(
            {
                "slide_index": 1,
                "shape_name_or_index": "Diagram",
                "action": action,
                field: 2,
            }
        )
    )
    app = smartart_session.get_application.return_value
    native_options = getattr(app, collection)
    native_options.assert_called_once_with(2)
    shapes = smartart_session.get_document.return_value.Slides().Shapes
    diagram = shapes.AddSmartArt.return_value.SmartArt
    assert getattr(diagram, attribute) is native_options.return_value
    assert result == {"success": True, "action": action, "total_nodes": None}


def test_non_macos_smartart_node_text_path_is_preserved(
    smartart_session: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    diagram = Mock(AllNodes=Mock(Count=1))
    shape = Mock(Type=smartart.MSO_SMART_ART, SmartArt=diagram)
    monkeypatch.setattr(smartart, "_get_shape", Mock(return_value=shape))
    result = smartart.modify_smartart(
        smartart.ModifySmartArtInput(
            slide_index=1,
            shape_name_or_index="Diagram",
            action="set_text",
            node_index=1,
            text="Changed",
        )
    )
    assert diagram.AllNodes.return_value.TextFrame2.TextRange.Text == "Changed"
    assert result["success"] and result["total_nodes"] == 1


@pytest.mark.parametrize(
    "operation", ["start", "stop", "next", "previous", "goto", "get_status"]
)
def test_macos_slideshow_refuses_unsafe_native_lifecycle(
    monkeypatch: pytest.MonkeyPatch, operation: str
) -> None:
    presentation = Mock(spec=JsObject)
    session = Mock()
    session.get_document.return_value = presentation
    presentation.Slides = Mock(Count=1)
    monkeypatch.setattr(slideshow, "get_session", Mock(return_value=session))
    params = (
        slideshow.SlideShowStartInput()
        if operation == "start"
        else slideshow.SlideShowGotoInput(slide_index=1)
        if operation == "goto"
        else DocumentInput()
    )
    with pytest.raises(NotImplementedError, match="slide show.*unsafe"):
        getattr(slideshow, "slideshow_" + operation)(params)
    assert presentation.mock_calls == []
    session.get_application.assert_not_called()
