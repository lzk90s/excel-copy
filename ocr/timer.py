from threading import Timer


class RepeatingTimer(Timer):
    def run(self):
        while not self.finished.is_set():
            self.function(*self.args, **self.kwargs)
            self.finished.wait(self.interval)


class UseTimer:
    def __init__(self, interval, function_name, *args, **kwargs):
        """
        :param interval:时间间隔
        :param function_name:可调用的对象
        :param args:args和kwargs作为function_name的参数
        """
        self.timer = RepeatingTimer(interval, function_name, *args, **kwargs)

    def timer_start(self):
        self.timer.start()

    def timer_cancle(self):
        self.timer.cancel()
