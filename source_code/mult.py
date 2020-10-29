"""
在 Python 中使用多处理和线程的示例。
此示例创建一个线程和一个进程，每个进程都调用相同的
函数。
"""

import multiprocessing as mp
import threading as td


def job(a, d):
    """
    此函数由线程和进程调用。
    它接受两个参数 a 和 d，并打印一条消息。
    """
    print('aaaaa')


if __name__ == '__main__':
    # 创建一个线程和一个进程，每个进程都使用参数 （1,2） 调用 job
    t1 = td.Thread(target=job, args=(1, 2))
    p1 = mp.Process(target=job, args=(1, 2))
    # 启动线程和进程
    t1.start()
    p1.start()
    # 等待线程和进程完成
    t1.join()
    p1.join()
