def normalize_url(url_):
    """规范化 URL"""
    return url_ if not url_ or url_.startswith(('http://', 'https://')) else 'https://' + url_


if __name__ == '__main__':
    print(normalize_url('zk.sceea.cn/'))
