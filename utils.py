import argparse
import datetime


def arg_parser():
    """arguments"""
    parser = argparse.ArgumentParser(description="自动将背景图片排列成需要的格式")
    parser.add_argument('-n', '--number', default=2, type=int, help="每页上放置的图片数量 n*n")
    parser.add_argument('-m', '--margin', default=0.5, type=float, help="页边距 cm")
    args = parser.parse_args()
    return args


def tid_maker():
	return '{0:%Y%m%d%H%M%S%f}'.format(datetime.datetime.now())
