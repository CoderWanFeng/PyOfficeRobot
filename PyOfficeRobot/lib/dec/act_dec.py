from functools import wraps


# 统一的动作输出
def act_info(act='消息'):
    def wrapper(func):
        @wraps(func)
        def inner_wrapper(*args, **kwargs):
            print(f'开始：发送{act}')
            res = func(*args, **kwargs)
            print(f'结束：发送{act}')
            return res

        return inner_wrapper

    return wrapper
