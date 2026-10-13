import datetime as dt

import pytest
from requests import Response
from requests.exceptions import HTTPError

from O365 import Account
from O365.connection import MSGraphProtocol


class MockResponse:
    def __init__(self, payload, status_code=200):
        self._payload = payload
        self.status_code = status_code

    def json(self):
        return self._payload

    def __bool__(self):
        return True


class MockConnection:
    def __init__(self):
        self.url = None
        self.params = None
        self.payload = {"value": []}
        self.error = None

    def get(self, url, params=None, **kwargs):
        self.url = url
        self.params = params
        if self.error:
            raise self.error
        return MockResponse(self.payload)


def _build_http_error(status_code):
    response = Response()
    response.status_code = status_code
    return HTTPError(f"{status_code} error", response=response)


@pytest.fixture
def work_hours():
    account = Account(
        ("client id", "client secret"),
        protocol=MSGraphProtocol(timezone="UTC"),
    )
    work_hours = account.work_hours_and_locations()
    work_hours.con = MockConnection()
    return work_hours


def test_occurrences_view_builds_me_url(work_hours):
    start = dt.datetime(2026, 1, 1, 9, 0, tzinfo=dt.timezone.utc)
    end = dt.datetime(2026, 1, 1, 17, 0, tzinfo=dt.timezone.utc)

    work_hours.occurrences_view(start, end)

    assert (
        work_hours.con.url
        == "https://graph.microsoft.com/v1.0/me/settings/workHoursAndLocations/"
        "occurrencesView(startDateTime='2026-01-01T09:00:00Z',endDateTime='2026-01-01T17:00:00Z')"
    )


def test_occurrences_view_builds_users_url_and_query(work_hours):
    start = dt.datetime(2026, 1, 1, 9, 0, tzinfo=dt.timezone.utc)
    end = dt.datetime(2026, 1, 1, 17, 0, tzinfo=dt.timezone.utc)

    work_hours.occurrences_view(
        start, end, user="alice@example.com", top=5, select=["id", "startDateTime"]
    )

    assert (
        work_hours.con.url
        == "https://graph.microsoft.com/v1.0/users/alice@example.com/settings/workHoursAndLocations/"
        "occurrencesView(startDateTime='2026-01-01T09:00:00Z',endDateTime='2026-01-01T17:00:00Z')"
    )
    assert work_hours.con.params == {"$top": 5, "$select": "id,startDateTime"}


def test_occurrence_parses_aware_datetime(work_hours):
    work_hours.con.payload = {
        "value": [
            {
                "id": "occ-1",
                "recurrenceId": "rec-1",
                "workLocationType": "office",
                "startDateTime": {"dateTime": "2026-04-03T10:00:00+02:00"},
                "endDateTime": {"dateTime": "2026-04-03T11:00:00+02:00"},
            }
        ]
    }

    occurrences = work_hours.occurrences_view(
        dt.datetime(2026, 4, 3, 0, 0, tzinfo=dt.timezone.utc),
        dt.datetime(2026, 4, 4, 0, 0, tzinfo=dt.timezone.utc),
    )

    assert len(occurrences) == 1
    assert occurrences[0].start == dt.datetime(
        2026, 4, 3, 8, 0, tzinfo=dt.timezone.utc
    )
    assert occurrences[0].end == dt.datetime(2026, 4, 3, 9, 0, tzinfo=dt.timezone.utc)


def test_occurrence_parses_naive_datetime_with_iana_timezone(work_hours):
    work_hours.con.payload = {
        "value": [
            {
                "id": "occ-2",
                "workLocationType": "remote",
                "startDateTime": {
                    "dateTime": "2026-01-03T09:00:00",
                    "timeZone": "America/Los_Angeles",
                },
                "endDateTime": {
                    "dateTime": "2026-01-03T10:00:00",
                    "timeZone": "America/Los_Angeles",
                },
            }
        ]
    }

    occurrences = work_hours.occurrences_view(
        dt.datetime(2026, 1, 3, 0, 0, tzinfo=dt.timezone.utc),
        dt.datetime(2026, 1, 4, 0, 0, tzinfo=dt.timezone.utc),
    )

    assert occurrences[0].start == dt.datetime(
        2026, 1, 3, 17, 0, tzinfo=dt.timezone.utc
    )
    assert occurrences[0].end == dt.datetime(2026, 1, 3, 18, 0, tzinfo=dt.timezone.utc)


def test_occurrence_parses_naive_datetime_with_windows_timezone(work_hours):
    work_hours.con.payload = {
        "value": [
            {
                "id": "occ-3",
                "workLocationType": "office",
                "startDateTime": {
                    "dateTime": "2026-01-03T09:00:00",
                    "timeZone": "Pacific Standard Time",
                },
                "endDateTime": {
                    "dateTime": "2026-01-03T10:00:00",
                    "timeZone": "Pacific Standard Time",
                },
            }
        ]
    }

    occurrences = work_hours.occurrences_view(
        dt.datetime(2026, 1, 3, 0, 0, tzinfo=dt.timezone.utc),
        dt.datetime(2026, 1, 4, 0, 0, tzinfo=dt.timezone.utc),
    )

    assert occurrences[0].start == dt.datetime(
        2026, 1, 3, 17, 0, tzinfo=dt.timezone.utc
    )
    assert occurrences[0].end == dt.datetime(2026, 1, 3, 18, 0, tzinfo=dt.timezone.utc)
    assert occurrences[0].time_zone == "Pacific Standard Time"


@pytest.mark.parametrize("status_code", [403, 404])
def test_occurrences_view_unavailable_endpoint_error_has_context(
    work_hours, status_code
):
    work_hours.con.error = _build_http_error(status_code)

    with pytest.raises(HTTPError) as exc:
        work_hours.occurrences_view(
            dt.datetime(2026, 1, 1, 0, 0, tzinfo=dt.timezone.utc),
            dt.datetime(2026, 1, 2, 0, 0, tzinfo=dt.timezone.utc),
        )

    message = str(exc.value)
    assert "unavailable" in message
    assert "Endpoint:" in message
    assert "workHoursAndLocations/occurrencesView" in message


def test_occurrences_view_transient_error_has_context(work_hours):
    work_hours.con.error = _build_http_error(500)

    with pytest.raises(HTTPError) as exc:
        work_hours.occurrences_view(
            dt.datetime(2026, 1, 1, 0, 0, tzinfo=dt.timezone.utc),
            dt.datetime(2026, 1, 2, 0, 0, tzinfo=dt.timezone.utc),
        )

    message = str(exc.value)
    assert "transient" in message
    assert "Endpoint:" in message


def test_occurrences_view_malformed_payload_raises(work_hours):
    work_hours.con.payload = {"unexpected": []}

    with pytest.raises(ValueError):
        work_hours.occurrences_view(
            dt.datetime(2026, 1, 1, 0, 0, tzinfo=dt.timezone.utc),
            dt.datetime(2026, 1, 2, 0, 0, tzinfo=dt.timezone.utc),
        )


def test_occurrences_view_malformed_occurrence_raises(work_hours):
    work_hours.con.payload = {"value": [{"workLocationType": "office"}]}

    with pytest.raises(ValueError):
        work_hours.occurrences_view(
            dt.datetime(2026, 1, 1, 0, 0, tzinfo=dt.timezone.utc),
            dt.datetime(2026, 1, 2, 0, 0, tzinfo=dt.timezone.utc),
        )
