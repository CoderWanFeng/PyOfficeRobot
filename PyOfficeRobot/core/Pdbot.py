import sqlite3


class BRAIN:
    """机器人的大脑： /.putin /.find """
    def __init__(self, path, tup):
        self.lie_name = tup
        self.path = path
        self.conn = sqlite3.connect(self.path)
        self.c = self.conn.cursor()
        try:
            self.c.execute(f'''CREATE TABLE BRAIN({self.lie_name})''')
        except sqlite3.OperationalError:
            pass

    def __open__(self):
        """打开数据库"""
        self.conn = sqlite3.connect(self.path)
        self.c = self.conn.cursor()

    def __cmd__(self, command):
        """数据库命令"""
        self.__open__()
        rang = self.c.execute(f"{command}")
        for row in rang:
            return row[1]

    def putin(self, msg):
        """向数据库加数据
        :type msg: str,tuple,list[tuple]
        """
        self.__open__()
        rang = self.all_question
        if type(msg) == str and '+' in msg:
            msgs = msg.split("+")
            if msgs[0] in self.sp_question[0].keys() or msgs[0] in self.sp_question[1].keys():
                self.delete(msgs[0])
            ds = (msgs[0], msgs[1], 'lang') if len(msg.split("+")) == 2 \
                else (tuple(msgs) if len(msgs) == 3 else ('测试', '这是个测试！', 'lang'))
            self.c.execute(f"INSERT INTO BRAIN VALUES (?,?,?)", ds) if ds not in rang else print('数据已有！[str]')
        elif type(msg) == tuple:
            ds = msg
            self.c.execute(f"INSERT INTO BRAIN VALUES (?,?,?)", ds) if ds not in rang else print('数据已有！[tuple]')
        elif type(msg) == list:
            ds = msg
            for dl in ds:
                print('数据已有！[list]') if dl in rang else print()
            self.c.executemany(f"INSERT INTO BRAIN VALUES (?,?,?)", msg)
        else:
            ds = list(('测试', '失败'))
        self.conn.commit()
        self.conn.close()
        return f'{ds[0]}：{ds[1]}'

    def find(self, item: str, tup: str) -> list[tuple]:
        """查找具体的项目，3个表头均可"""
        self.__open__()
        rang = self.c.execute(f'SELECT * FROM BRAIN WHERE {item}=?', (tup,))
        return [raw for raw in rang]

    def love(self, xinxi):
        """亲密问题"""
        self.__open__()
        return self.__cmd__(f"SELECT * FROM BRAIN WHERE question='{xinxi}' AND type='love'")

    def lang(self, xinxi):
        """普通问题"""
        self.__open__()
        return self.__cmd__(f"SELECT * FROM BRAIN WHERE question='{xinxi}' AND type='lang'")

    def all(self, xinxi):
        """所有问题"""
        self.__open__()
        return self.__cmd__(f"SELECT * FROM BRAIN WHERE question='{xinxi}'")

    def delete(self, question, typename=''):
        """删除某个问题"""
        self.__open__()
        if typename == '':
            self.c.execute(f"DELETE FROM BRAIN WHERE question='{question}'")
        else:
            self.c.execute(f"DELETE FROM BRAIN WHERE question='{question}' AND type='{typename}'")
        self.conn.commit()
        self.conn.close()
        self.__open__()
        pr = f'[{question}]问题删除！' if typename == '' else f'[{question}|{typename}]问题删除！'
        return pr

    def update(self, question, answer=''):
        """将问题升级或者回复修改"""
        self.__open__()
        if answer == '':
            self.c.execute(f"UPDATE BRAIN SET type = 'love' WHERE question = '{question}'")
        else:
            self.c.execute(f"UPDATE BRAIN SET answer = '{answer}' WHERE question = '{question}'")
        self.conn.commit()
        self.conn.close()
        self.conn = sqlite3.connect(self.path)
        return f'[{question}]问题升级！'

    @property
    def all_question(self) -> list:
        """快速显示所有的问题"""
        self.__open__()
        rang = self.c.execute(f'SELECT * FROM BRAIN')
        return [raw[0] for raw in rang]

    @property
    def sp_question(self):
        """显示所有的问题，将私密问题和公开问题分开"""
        dic_love, dic_lang = {}, {}
        for raw in self.findall:
            if raw[2] == 'love':
                dic_love[raw[0]] = raw[1]
            elif raw[2] == 'lang':
                dic_lang[raw[0]] = raw[1]
        return [dic_love, dic_lang]

    @property
    def all_sheet(self) -> list:
        """返回所有表头"""
        self.__open__()
        return [x for x in self.c.execute("SELECT tbl_name FROM sqlite_master")]

    @property
    def findall(self):
        """找到所有问题"""
        self.__open__()
        rang = self.c.execute(f'SELECT * FROM BRAIN')
        return rang


def main():
    """这是个测试函数！"""
    data = BRAIN(f'I:/EXE/WX/Log/Data_brain.db', 'question,answer,type')
    # for i in data.findall:
    #     print(i)
    print(data.find('type', 'lang'))


if __name__ == '__main__':
    """该函数仅在本模块儿运行"""
    main()
