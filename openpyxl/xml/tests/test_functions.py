import pytest

from openpyxl.xml.functions import ConditionalElement


@pytest.fixture
def root():
    from openpyxl.xml.functions import Element
    return Element("root")


@pytest.mark.parametrize("condition", [True, 1, -1])
def test_simple(root, condition):
    ConditionalElement(root, "start", condition)
    assert root.find("start").tag == "start"


def test_simple_attrib(root):
    ConditionalElement(root, "start", True, 'val')
    tag = root.find("start")
    assert tag.attrib == {'val': '1'}


def test_dict_attrib(root):
    ConditionalElement(root, "start", True, {'val':'single'})
    tag = root.find("start")
    assert tag.attrib == {'val':'single'}


@pytest.mark.parametrize("condition", [False, 0, None])
def test_no_tag(root, condition):
    ConditionalElement(root, "start", condition)
    assert root.find("start") is None
