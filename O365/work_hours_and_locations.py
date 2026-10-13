import datetime as dt
import logging
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from dateutil.parser import parse
from requests.exceptions import HTTPError

from .utils import ApiComponent, get_iana_tz

log = logging.getLogger(__name__)


class WorkPlanOccurrence(ApiComponent):
    """A work-hours and location occurrence returned by Microsoft Graph."""

    def __init__(self, *, parent=None, con=None, **kwargs):
        if parent and con:
            raise ValueError("Need a parent or a connection but not both")
        self.con = parent.con if parent else con

        main_resource = kwargs.pop("main_resource", None) or (
            getattr(parent, "main_resource", None) if parent else None
        )

        super().__init__(
            protocol=parent.protocol if parent else kwargs.get("protocol"),
            main_resource=main_resource,
        )

        cloud_data = kwargs.get(self._cloud_data_key, {})
        cc = self._cc

        self.object_id = cloud_data.get(cc("id"))
        if not self.object_id:
            raise ValueError("Missing occurrence id")
        self.id = self.object_id

        self.recurrence_id = cloud_data.get(cc("recurrenceId"))
        self.work_location_type = cloud_data.get(cc("workLocationType"), "")

        start_payload = cloud_data.get(cc("startDateTime")) or cloud_data.get(
            cc("start")
        )
        end_payload = cloud_data.get(cc("endDateTime")) or cloud_data.get(cc("end"))

        self.time_zone = (
            cloud_data.get(cc("timeZone"))
            or self._extract_time_zone(start_payload)
            or self._extract_time_zone(end_payload)
        )

        self.start = self._parse_occurrence_datetime(start_payload, self.time_zone)
        self.end = self._parse_occurrence_datetime(end_payload, self.time_zone)

        self.location = cloud_data.get(cc("location")) or cloud_data.get(
            cc("workLocation")
        )

        if self.start is None or self.end is None:
            raise ValueError("Occurrence is missing a valid start or end datetime")

    def _extract_time_zone(self, payload):
        if not isinstance(payload, dict):
            return None
        return payload.get(self._cc("timeZone"))

    def _resolve_time_zone(self, time_zone_name):
        if time_zone_name:
            try:
                return ZoneInfo(time_zone_name)
            except ZoneInfoNotFoundError:
                try:
                    return get_iana_tz(time_zone_name)
                except ZoneInfoNotFoundError:
                    log.debug(
                        "TimeZone not found in IANA or Windows maps: %s. Using protocol timezone.",
                        time_zone_name,
                    )

        return self.protocol.timezone

    def _parse_occurrence_datetime(self, payload, fallback_time_zone):
        if payload is None:
            return None

        cc = self._cc
        date_time_raw = payload
        time_zone_name = fallback_time_zone

        if isinstance(payload, dict):
            date_time_raw = payload.get(cc("dateTime"))
            time_zone_name = payload.get(cc("timeZone")) or fallback_time_zone

        if not date_time_raw:
            return None

        parsed = parse(date_time_raw)

        if parsed.tzinfo is not None:
            date_time = parsed
        else:
            date_time = parsed.replace(tzinfo=self._resolve_time_zone(time_zone_name))

        return date_time.astimezone(self.protocol.timezone)


class WorkHoursAndLocations(ApiComponent):
    """Read work-hours and location occurrences from Microsoft Graph."""

    occurrence_constructor = WorkPlanOccurrence  #: :meta private:

    def __init__(self, *, parent=None, con=None, **kwargs):
        if parent and con:
            raise ValueError("Need a parent or a connection but not both")
        self.con = parent.con if parent else con

        main_resource = kwargs.pop("main_resource", None) or (
            getattr(parent, "main_resource", None) if parent else None
        )

        super().__init__(
            protocol=parent.protocol if parent else kwargs.get("protocol"),
            main_resource=main_resource,
        )

    @staticmethod
    def _validate_datetime(value, name):
        if not isinstance(value, dt.datetime):
            raise ValueError(f"'{name}' must be a datetime instance")
        return value

    def _serialize_graph_datetime(self, value):
        if value.tzinfo is None:
            value = value.replace(tzinfo=self.protocol.timezone)
        return value.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def _build_occurrences_view_url(self, start_datetime, end_datetime, user=None):
        resource = "/me" if user is None else f"/users/{user}"
        start_value = self._serialize_graph_datetime(start_datetime)
        end_value = self._serialize_graph_datetime(end_datetime)
        endpoint = (
            "/settings/workHoursAndLocations/occurrencesView"
            f"(startDateTime='{start_value}',endDateTime='{end_value}')"
        )
        return f"{self.protocol.service_url.rstrip('/')}{resource}{endpoint}"

    def occurrences_view(
        self, start_datetime, end_datetime, user=None, top=None, select=None
    ):
        """Return work plan occurrences in a datetime range.

        :param datetime start_datetime: range start datetime
        :param datetime end_datetime: range end datetime
        :param str user: optional user id/upn, defaults to /me when omitted
        :param int top: optional OData $top
        :param list[str] select: optional OData $select fields
        :rtype: list[WorkPlanOccurrence]
        """
        start_datetime = self._validate_datetime(start_datetime, "start_datetime")
        end_datetime = self._validate_datetime(end_datetime, "end_datetime")

        if end_datetime <= start_datetime:
            raise ValueError("'end_datetime' must be greater than 'start_datetime'")

        if user is not None and not isinstance(user, str):
            raise ValueError("'user' must be a string when provided")

        url = self._build_occurrences_view_url(start_datetime, end_datetime, user=user)

        params = {}
        if top is not None:
            if not isinstance(top, int) or top <= 0:
                raise ValueError("'top' must be a positive integer")
            params["$top"] = top
        if select is not None:
            if not isinstance(select, (list, tuple)) or not all(
                isinstance(field, str) for field in select
            ):
                raise ValueError("'select' must be a list of strings")
            params["$select"] = ",".join(select)

        try:
            response = self.con.get(url, params=params or None)
        except HTTPError as e:
            status_code = e.response.status_code if e.response else None
            if status_code in (403, 404):
                message = (
                    "Work-hours endpoint is unavailable for this tenant/user "
                    f"(status: {status_code})"
                )
            else:
                message = (
                    "Failed to fetch work-hours occurrences; this may be transient"
                )
                if status_code:
                    message = f"{message} (status: {status_code})"

            raise HTTPError(f"{message} | Endpoint: {url}", response=e.response) from e

        if not response:
            return []

        payload = response.json()
        values = payload.get(self._cc("value"), payload.get("value"))

        if values is None:
            raise ValueError(
                "Malformed response payload: missing 'value' collection for occurrences view"
            )
        if not isinstance(values, list):
            raise ValueError("Malformed response payload: 'value' must be a list")

        occurrences = []
        for index, item in enumerate(values):
            if not isinstance(item, dict):
                raise ValueError(
                    f"Malformed response payload: occurrence at index {index} is not an object"
                )

            try:
                occurrences.append(
                    self.occurrence_constructor(parent=self, **{self._cloud_data_key: item})
                )
            except Exception as e:
                raise ValueError(
                    f"Malformed occurrence payload at index {index}: {e}"
                ) from e

        return occurrences
