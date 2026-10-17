"""Element-ID validation: SAP screen names in path segments ('#', '*', '=>', '<FS>')."""

import pytest

from mcp_sap_gui.controller import _ELEMENT_ID_RE, SAPGUIControllerBase

SO23_DELETE = (
    "wnd[0]/usr/tabsSO23_TAB/tabpTAB1/ssubSUB1:SAPLSO14:0960/btn%#AUTOTEXT005"
)

ACCEPTED = [
    SO23_DELETE,
    "wnd[0]/usr/subSUB1:SAPLSO14:0960/btn%#AUTOTEXT005",
    "wnd[1]/usr/btn%#AUTOTEXT001",
    "wnd[0]/usr/tblX/btn%#AUTOTEXT005[0,1]",
    # Names read from live screens and the screen texts table D021T.
    "wnd[0]/usr/lbl*DY_1000-CLIF",
    "wnd[0]/usr/radDYNPROLOGIC=>DYNPFIELDS-SELECTED_CHECKS",
    "wnd[0]/usr/txtGO_REPORT->MV_NUMBER_OF_FREE_JOBS",
    "wnd[0]/usr/txt<FILTER>-FIELD_LABEL",
    "wnd[0]/usr/chk=/BA1/V_TF4_FXRTY-XFIXD",
    "wnd[0]/usr/btnB_P++",
    "wnd[0]/usr/txt&LOCFIELD0001",
    "wnd[0]/usr/lbl#%'AUTOTEXT01",
    "wnd[0]/usr/btnFE_(",
    # Shapes accepted before the change stay accepted.
    "wnd[0]/usr/txtTHOLT-KTEXT",
    "wnd[0]/shellcont[0]/shell",
    "wnd[0]/tbar[0]/btn[11]",
]
REJECTED = [
    "wnd[0]/foo/btn%#X",  # non-standard first segment
    "wnd[0]/usr/btn[#]",  # '#' inside an index
    "wnd[0]/usr/txt#[0",  # unbalanced bracket
    "usr/btn%#AUTOTEXT005",  # no window
    "wnd[0]/usr/txt A",  # whitespace
    "wnd[0]/usr/txtA\tB",
    "wnd[0]/usr/txtA]",  # stray bracket
    "wnd[0]/usr//txtA",  # empty segment
]


@pytest.mark.parametrize("element_id", ACCEPTED)
def test_accepted(element_id: str) -> None:
    assert _ELEMENT_ID_RE.fullmatch(element_id)


@pytest.mark.parametrize("element_id", REJECTED)
def test_rejected(element_id: str) -> None:
    assert not _ELEMENT_ID_RE.fullmatch(element_id)


class _FakeSession:
    def __init__(self) -> None:
        self.found: list = []

    def findById(self, element_id: str) -> str:  # noqa: N802 - COM name
        self.found.append(element_id)
        return element_id


def _controller() -> SAPGUIControllerBase:
    # Skip __init__: it imports pywin32, which these tests do not need.
    controller = SAPGUIControllerBase.__new__(SAPGUIControllerBase)
    controller._session = _FakeSession()
    return controller


def test_find_element_resolves_autotext_button() -> None:
    controller = _controller()
    full_path = "/app/con[0]/ses[0]/" + SO23_DELETE
    assert controller._find_element(full_path) == SO23_DELETE
    assert controller._session.found == [SO23_DELETE]


@pytest.mark.parametrize("element_id", REJECTED)
def test_find_element_still_rejects(element_id: str) -> None:
    controller = _controller()
    with pytest.raises(ValueError, match="Invalid SAP element ID"):
        controller._find_element(element_id)
    assert controller._session.found == []
