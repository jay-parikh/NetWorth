"""Auto-update check — version comparison and hint logic (no real network)."""

from networth.update import _parse_version, check_for_update


class _FakeResp:
    def __init__(self, status, payload):
        self.status_code = status
        self._payload = payload

    def json(self):
        return self._payload


class _FakeSession:
    def __init__(self, status=200, payload=None):
        self._resp = _FakeResp(status, payload or {})

    def get(self, *a, **k):
        return self._resp


def test_version_ordering():
    assert _parse_version("v1.1.0") > _parse_version("v1.0.0")
    assert _parse_version("1.2.0") > _parse_version("1.1.9")
    # a final release beats its own release candidates
    assert _parse_version("1.1.0") > _parse_version("1.1.0rc3")
    assert _parse_version("1.1.0rc3") > _parse_version("1.1.0rc1")
    assert _parse_version("v1.1.0-rc.2") == _parse_version("1.1.0rc2")
    assert _parse_version("garbage") is None


def test_hint_when_newer_release_exists():
    sess = _FakeSession(200, {"tag_name": "v1.2.0",
                              "html_url": "https://example/releases/v1.2.0"})
    hint = check_for_update("1.1.0", session=sess)
    assert hint and "v1.2.0" in hint and "v1.1.0" in hint


def test_no_hint_when_same_or_older():
    same = _FakeSession(200, {"tag_name": "v1.1.0"})
    assert check_for_update("1.1.0", session=same) is None
    older = _FakeSession(200, {"tag_name": "v1.0.0"})
    assert check_for_update("1.1.0", session=older) is None


def test_running_rc_sees_the_final_release():
    sess = _FakeSession(200, {"tag_name": "v1.1.0"})
    assert check_for_update("1.1.0rc3", session=sess) is not None


def test_silent_on_error_status_or_no_releases():
    assert check_for_update("1.1.0", session=_FakeSession(404, {})) is None
    assert check_for_update("1.1.0", session=_FakeSession(200, {})) is None  # no tag


def test_never_raises_on_network_failure():
    class Boom:
        def get(self, *a, **k):
            raise RuntimeError("network down")
    assert check_for_update("1.1.0", session=Boom()) is None
