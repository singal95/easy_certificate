import argparse
import datetime


def arg_parser():
    """arguments"""
    parser = argparse.ArgumentParser(description="自动将背景图片排列成需要的格式")
    parser.add_argument('--rows', default=2, type=int, help="每页多少行图片")
    parser.add_argument('--collumns', default=1, type=int, help="每页多少列图片")
    parser.add_argument('--margin', default=0.5, type=float, help="页边距 cm")
    parser.add_argument('--orientation', default='vertical', type=str, help='文本方向')
    args = parser.parse_args()
    return args


def tid_maker():
	return '{0:%Y%m%d%H%M%S%f}'.format(datetime.datetime.now())
