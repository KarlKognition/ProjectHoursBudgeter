"""Testiong of Explanation Page"""
import sys
import types
from collections.abc import Iterator
import pytest
from PyQt6.QtWidgets import QApplication, QLabel, QWizardPage, QHBoxLayout

@pytest.fixture(autouse=True)
def mock_page_utils(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    """
    Mock phb_app.utils.page_utils to isolate ExplanationPage
    from the rest of the application.
    """
    fake_pu = types.SimpleNamespace()

    def fake_set_titles(page: QWizardPage, title: str, subtitle: str) -> None:
        """Set title and subtitle directly on the provided page."""
        page.setTitle(title)
        page.setSubTitle(subtitle)
    
    def fake_create_watermark_label(_images_dir: str, _fail_msg: str) -> QLabel:
        """Return a trivial QLabel standing in for the watermark widget."""
        return QLabel("Watermark")
    
    def fake_create_intro_message(text: str) -> QLabel:
        """Return a QLabel containing the provided intro text."""
        return QLabel(text)
    
    def fake_setup_page(
        page: QWizardPage,
        widgets: list[QLabel],
        layout: QHBoxLayout,
        *,
        spacing: int,
        margins: tuple[ int, int, int, int]
    ) -> None:
        """Attach widgets and layout to page."""
        for w in widgets:
            layout.addWidget(w)
        layout.setSpacing(spacing)
        layout.setContentsMargins(*margins)
        page.setLayout(layout)

    fake_pu.set_titles = fake_set_titles
    fake_pu.create_watermark_label = fake_create_watermark_label
    fake_pu.create_intro_message = fake_create_intro_message
    fake_pu.setup_page = fake_setup_page

    monkeypatch.setitem(sys.modules, "phb_app.utils.page_utils", fake_pu)
    yield


def test_explanation_page_inherits_qwizardpage(qapp: QApplication):
    """
    Verify that ExplanationPage is a subclass of QWizardPage.
    """
    from phb_app.wizard.pages.explanation import ExplanationPage # pylint: disable=import-outside-toplevel

    page:QWizardPage = ExplanationPage()
    qapp.processEvents()
    assert isinstance(page,QWizardPage)


def test_explanation_page_sets_titles(qapp: QApplication) -> None:
    """Ensure ExplanationPage sets non-empty title and subtitle."""
    from phb_app.wizard.pages.explanation import ExplanationPage # pylint: disable=import-outside-toplevel

    page: QWizardPage = ExplanationPage()
    qapp.processEvents()
    assert page.title().strip() != ""
    assert page.subTitle().strip() != ""

def test_explanation_page_layout_properties(qapp: QApplication) -> None:
    """Ensure ExplanationPage uses QHBoxLayout with correct margins and spacing."""
    from phb_app.wizard.pages.explanation import ExplanationPage # pylint: disable=import-outside-toplevel

    page: QWizardPage = ExplanationPage()
    qapp.processEvents()
    layout = page.layout()
    assert isinstance(layout, QHBoxLayout)

    margins = layout.contentsMargins()
    assert (margins.left(), margins.top(), margins.right(), margins.bottom()) == (25, 25, 25, 25)
    assert layout.spacing() == 35

def test_explanation_page_has_widgets(qapp: QApplication) -> None:
    """Ensure ExplanationPage inserts at least two QLabel widgets into its layout."""
    from phb_app.wizard.pages.explanation import ExplanationPage # pylint: disable=import-outside-toplevel

    page: QWizardPage = ExplanationPage()
    qapp.processEvents()
    layout = page.layout()
    count = layout.count()
    assert count >= 2

    widgets = [layout.itemAt(i).widget() for i in range(count)]
    assert all(isinstance(w, QLabel) for w in widgets)
