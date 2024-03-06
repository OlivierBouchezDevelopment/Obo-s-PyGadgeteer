import json
import datetime

## FROM : https://stackoverflow.com/questions/12122007/python-json-encoder-to-support-datetime


class DateTimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime.datetime, datetime.date, datetime.time)):
            return obj.isoformat()
        elif isinstance(obj, datetime.timedelta):
            return (datetime.datetime.min + obj).time().isoformat()

        return super(DateTimeEncoder, self).default(obj)


if __name__ == "__main__":
    now = datetime.datetime.now()
    encoder = DateTimeEncoder()
    data = {"datetime": now, "date": now.date(), "time": now.time()}

    print(json.dumps(data, cls=DateTimeEncoder))
