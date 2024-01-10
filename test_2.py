import datetime
import http.client

from dateutil.parser import parse


def get_webserver_time() -> datetime:
    """获取网络时间，如果网络异常则返回None"""
    try:
        time_conn = http.client.HTTPConnection('www.baidu.com')
        time_conn.request("GET", "/")
        r = time_conn.getresponse()
        ts = r.getheader('date')  # 获取http头date部分
        print(ts)
        ltime = parse(ts).date()  # 将GMT时间转换成北京时间
        return ltime
    except Exception as e:
        print(e)
        print('获取网络时间失败！')
        return None


def get_now_date_time():
    """获取当前日期和时间"""
    # 获取当前日期和时间
    now_date_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # 将当前的时间和日期加10分钟
    new_date_time = parse(now_date_time) + datetime.timedelta(minutes=10)
    # time = parse(now_date_time)
    # 将dateTimeEdit的日期和时间设置为当前日期和时间
    # self.dateTimeEdit.setDateTime(parse(now_date_time))


if __name__ == '__main__':
    # time_ = get_webserver_time()
    # print(time_)
    # print(type(time_))
    get_now_date_time()
