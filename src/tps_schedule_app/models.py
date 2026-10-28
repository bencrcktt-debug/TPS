from dataclasses import dataclass, field


TIME_RANGE_SEPARATOR = " \u2013 "
MINUTES_PER_DAY = 24 * 60


@dataclass(frozen=True)
class ScheduleEvent:
    description: str
    start_minutes: int
    end_minutes: int | None = None
    sheet_name: str = ""
    start_column: str = ""
    fill_id: int = 0

    @property
    def time_label(self) -> str:
        if self.end_minutes is None:
            return _format_clock(self.start_minutes, include_suffix=True)

        start_suffix = _suffix_for_minutes(self.start_minutes)
        end_suffix = _suffix_for_minutes(self.end_minutes)
        start_text = _format_clock(
            self.start_minutes,
            include_suffix=start_suffix != end_suffix,
        )
        end_text = _format_clock(self.end_minutes, include_suffix=True)
        return f"{start_text}{TIME_RANGE_SEPARATOR}{end_text}"


@dataclass(frozen=True)
class DaySchedule:
    sheet_name: str
    events: list[ScheduleEvent] = field(default_factory=list)


@dataclass(frozen=True)
class PersonSchedule:
    name: str
    days: list[DaySchedule] = field(default_factory=list)


@dataclass(frozen=True)
class WorkbookSchedule:
    sheet_names: list[str]
    people: list[PersonSchedule]


def _suffix_for_minutes(total_minutes: int) -> str:
    total_minutes %= MINUTES_PER_DAY
    return "am" if total_minutes < 12 * 60 else "pm"


def _format_clock(total_minutes: int, include_suffix: bool) -> str:
    total_minutes %= MINUTES_PER_DAY
    hour_24, minute = divmod(total_minutes, 60)
    hour_12 = hour_24 % 12 or 12
    text = f"{hour_12}:{minute:02d}"
    if include_suffix:
        return f"{text} {_suffix_for_minutes(total_minutes)}"
    return text
