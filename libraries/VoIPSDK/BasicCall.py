# -*- coding: utf-8 -*-
"""
description: 呼叫中心SDK测试,
"""
import datetime
import json
import os
import sys
import threading

import clr
from SDKWrapper import *
from robotremoteserver import RobotRemoteServer

clr.AddReference('SDKWrapper')


class BasicCall(object):
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def __init__(self):
        self.instance = None
        self.call_thread = None
        self.call_result = None
        self.path = r"./CallEvent.py"
        self.event_info_path = "{0}/event_info.txt".format(os.getcwd())

    def user_login(self, local_ip, queue_num, num, pass_wd,
                   server_ip, server_port, role=0, user_id=''):
        """
        @description: 坐席迁入(登录)接口
        :param local_ip: 当前pc的IP，用于事件通知
        :param queue_num: 坐席所在队列的队列名称，在服务器配置该坐席属于哪个呼叫队列，其中队列名称和队列号码有区别
        :param num: 坐席号码
        :param pass_wd: 坐席登录密码默认666666
        :param role: 坐席角色，传"0"就行
        :param server_ip: 服务器IP地址
        :param server_port: 服务器端口号(值为:18083)
        :param user_id: 用户id, 本次不涉及直接传空
        :return: 如果成功则返回”OK”，否则返回失败错误码。
        """
        try:
            result = self.instance.Login(str(local_ip), str(queue_num), str(num), str(pass_wd),
                                         str(role), str(server_ip), str(server_port), str(user_id))
            print(result)
            return result
        except BaseException as e:
            print("坐席登录异常：{0}".format(e))
            return False

    def int_instance(self):
        if not self.instance:
            self.instance = SDKProxy(self.path)
            print("instance content：{0}".format(self.instance))
            return self.instance

    def user_logout(self):
        """
        @description: 坐席退出登录
        :return:如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.Logout()
            print(result)
            return result
        except BaseException as e:
            print("坐席退出登录异常：{0}".format(e))

    def call_out(self, number):
        """
        @description: 外呼
        :param number: 外呼的号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        if self.call_thread:
            while True:
                if self.call_thread.is_alive():
                    print("call thread is alive")
                else:
                    print("call thread is finish:{0}".format(self.call_result))
                    print("last_call_res is:{0}".format(self.call_result))
                    break
        try:
            print("call_out_num is:{0}".format(number))
            self.call_thread = threading.Thread(
                target=self.instance.Call, args=(number,))
            self.call_result = self.call_thread.start()
            # result = self.instance.Call(str(number))
            print("*****call_res******:{0}".format(self.call_result))
            return self.call_result
        except BaseException as e:
            print("呼叫异常：{0}".format(e))

    def call_out_without_thread(self, number):
        try:
            print("call_out_num is:{0}".format(number))
            result = self.instance.Call(number)
            print(result)
            return result
        except BaseException as e:
            print("坐席退出登录异常：{0}".format(e))

    def hang_up(self):
        """
        @description: 挂断
        :return:如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.Hangup()
            print(result)
            return result
        except BaseException as e:
            print("挂断异常：{0}".format(e))

    def hold(self):
        """
        @description: 通话保持
        :return:如果执行成功的话返回”OK”,否则返回错误码。
        """
        try:
            result = self.instance.Hold()
            return result
        except BaseException as e:
            print("通话保持出现异常：{0}".format(e))

    def hold_over(self):
        """
        @description: 恢复通话
        :return:如果执行成功的话返回”OK”,否则返回错误码。
        """
        try:
            result = self.instance.HoldOver()
            return result
        except BaseException as e:
            print("恢复通话出现异常：{0}".format(e))

    def transfer(self, number):
        """
        @description: 盲转
        :param number: 转移的号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.Transfer(str(number))
            print("****盲转****：{0}".format(result))
            return result
        except BaseException as e:
            print("盲转出现异常：{0}".format(e))

    def consult(self, number, exp_result=''):
        """
        @description: 协商转
        :param number: 协商转移的号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            if not exp_result:
                call_thread = threading.Thread(
                    target=self.instance.Consult, args=(number,))
                self.call_result = call_thread.start()
            else:
                result = self.instance.Consult(str(number))
                print("****协商转consult****: {0}".format(result))
                return result
        except BaseException as e:
            print(u"协商转出现异常：{0}".format(e))

    def confirm_consult(self):
        """
        @description: 确认协商转
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.ConsultToAttXfer()
            print("****确认协商转****: {0}".format(result))
            return result
        except BaseException as e:
            print("协商转出现异常：{0}".format(e))

    def cancel_consult(self):
        """
        @description: 取消协商转
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.ConsultOver()
            print("****取消协商转****: {0}".format(result))
            return result
        except BaseException as e:
            print("协商转出现异常：{0}".format(e))

    def invite_target(self, target_num, exp_result=''):
        """
        @description: 邀请三方通话
        :param target_num: 目标号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            if not exp_result:
                call_thread = threading.Thread(
                    target=self.instance.ConsultToMeetingTarget, args=(target_num,))
                self.call_result = call_thread.start()
                return self.call_result
            else:
                result = self.instance.ConsultToMeetingTarget(str(target_num))
                print("****邀请三方通话****: {0}".format(result))
                return result
        except BaseException as e:
            print("邀请三方通话出现异常：{0}".format(e))

    def consult_switch(self):
        """
        @description: 轮流通话
        """
        try:
            result = self.instance.consultSwitch()
            return result
        except BaseException as e:
            print(u"轮流通话出现异常：{0}".format(e))

    def consult_to_meeting(self):
        """
        @description: 合并通话并形成会议
        """
        try:
            result = self.instance.consultToMeeting()
            return result
        except BaseException as e:
            print(u"合并通话并形成会议出现异常：{0}".format(e))

    def whisper(self, target_num):
        """
        @description: 耳语
        :param target_num: 目标号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.whisper(str(target_num))
            return result
        except BaseException as e:
            print("耳语出现异常：{0}".format(e))

    def monitor(self, number):
        """
        @description: 监听话务员
        :param number: 监听的号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.Monitor(str(number))
            return result
        except BaseException as e:
            print("邀请三方通话出现异常：{0}".format(e))

    def monitor_to_barge(self):
        """
        @description: 强插,只有在监听的状态下才能进行
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.MonitorToBarge()
            return result
        except BaseException as e:
            print("强插出现异常：{0}".format(e))

    def barge(self, number):
        """
        @description: 强插,通话中的状态下进行
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.Barge(str(number))
            return result
        except BaseException as e:
            print("强插出现异常：{0}".format(e))

    def force_release(self, number):
        """
        @description: 强拆,只有在监听的状况下才能进行
        :param number: 强拆的号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.ForceRelease(str(number))
            return result
        except BaseException as e:
            print(u"强拆出现异常：{0}".format(e))

    def monitor_to_break(self):
        """
        @description: 强替,只有在监听的状态下才能进行
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.MonitorToBreak()
            return result
        except BaseException as e:
            print(u"强替出现异常：{0}".format(e))

    def add_black_list(self, number, lockTime):
        """
        @description: 设置黑名单
        :param number: 要添加进黑名单号码
        :param lockTime: 黑名单锁定的时间
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.BlackListAddOne(str(number), str(lockTime))
            return result
        except BaseException as e:
            print(u"添加黑名单出现异常：{0}".format(e))

    def remove_black_list(self, number):
        """
        @description: 移除选中的黑名单
        :param number: 要移除的黑名单号码
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.BlackListRemoveOne(str(number))
            return result
        except BaseException as e:
            print(u"移除黑名单出现异常：{0}".format(e))

    def remove_all_black_list(self):
        """
        @description: 移除选中的黑名单
        :return: 如果执行成功的话返回”OK”,否则返回错误码
        """
        try:
            result = self.instance.BlackListRemoveAll()
            return result
        except BaseException as e:
            print(u"移除所有黑名单出现异常：{0}".format(e))

    def SubscribeCallList(self, queueNumber):
        """
        @description: 订阅当前队列的事件
        :param queueNumber: 要订阅队列的队列号码
        :return: 以OK@后面跟随json格式返回,如果订阅失败的话返回错误码。
        eg: “OK@[
             {
            "CallId":"1",
            "aNumber":"120",
            "bNumber":"1001",
            "time":"2018-10-19 14:28:13",
            "fromNumber":"1212"
             }]
        """
        try:
            result = self.instance.SubscribeCallList(str(queueNumber))
            return result
        except BaseException as e:
            print(u"订阅出现异常：{0}".format(e))

    def query_lostcall_count(self, queue_num, begin_time, end_time):
        """
        @description: 查询未接来电的数量
        :param queue_num: 队列名称
        :param begin_time: 开始时间
        :param end_time: 结束时间
        :return: 如果查询成功返回OK@加上值，否则返回错误码。例如:”OK@12”
        """
        try:
            result = self.instance.QueryLostCallsCount(
                str(queue_num), str(begin_time), str(end_time))
            return result
        except BaseException as e:
            print("查询未接来电总数出现异常：{0}".format(e))

    def query_lostcall_detail(
            self, page, page_size, queue_num, begin_time, end_time, order_by="desc"):
        """
        @description: 查询所有未接来电详细数据
        :param page: 当前第几页,0作为起始页
        :param page_size: 每页的大小
        :param queue_num: 队列号码
        :param begin_time: 开始时间
        :param end_time: 结束时间
        :return: 查询成功以OK@后面跟随json格式返回
        """

        key_words = {
            "queueName": queue_num,
            "beginTime": begin_time,
            "endTime": end_time,
            "orderBy": order_by
        }
        try:
            result = self.instance.QueryAllLostCall(
                int(page), int(page_size), json.dumps(key_words))
            return result
        except BaseException as e:
            print("查询未接来电详细数据异常：{0}".format(e))

    def query_harass_list(self, harassListName, page, pageSize):
        """
        @description: 查询骚扰列表
        :harassListName：列表名(暂时是table1和table2，后续改进)
        ：page：第几页
        ：pageSize：每页大小
        :return: 如果成功则返回”OK”，否则返回”-1”
        """
        try:
            result = self.instance.QueryHarassList(
                str(harassListName), int(page), int(pageSize))
            return result
        except BaseException as e:
            print("查询骚扰列表出现异常：{0}".format(e))

    def add_to_harass_list(self, harassListName, number, timeLen):
        """
        @description: 添加到骚扰列表
        :harassListName：列表名(暂时是table1和table2，后续改进)
        ：number：号码
        ：pageSize：时长(分钟)
        :return: 如果成功则返回”OK”，否则返回”-1”
        """
        try:
            result = self.instance.QueryHarassList(
                str(harassListName), str(number), int(timeLen))
            return result
        except BaseException as e:
            print("添加到骚扰列表出现异常：{0}".format(e))

    def delete_harass_list(self, numberList):
        """
        @description: 从骚扰列表中移除号码
        ：numberList：号码列表。例如[1000,1001,1002]
        :return: 如果成功则返回”OK”，否则返回”-1”
        """
        try:
            result = self.instance.DeleteHarassList(str(numberList))
            return result
        except BaseException as e:
            print("从骚扰列表中移除号码出现异常：{0}".format(e))

    def pick_up(self):
        """
        @description: 接听(摘机)
        :return: 如果成功则返回”OK”，否则返回”-1”
        """
        try:
            result = self.instance.PickUp()
            print("****接听****".format(result))
            return result
        except BaseException as e:
            print("摘机出现异常：{0}".format(e))

    def QueryStatus(self, number):
        """
        @description: 接听(摘机)
        :number：要查询的号码
        :return: 如果成功则返回”OK”，否则返回”-1”
        """
        try:
            result = self.instance.QueryStatus(number)
            return result
        except BaseException as e:
            print("查询状态出现异常：{0}".format(e))

    def QueryStatusList(self):
        """
        @description: 查询所有号码的状态
        :return: 如果成功则返回”OK”，否则返回”-1”
        """
        try:
            result = self.instance.QueryStatusList()
            return result
        except BaseException as e:
            print("查询状态出现异常：{0}".format(e))

    def AddToRedList(self, redLists):
        """
        @description: 添加红名单
        redLists为红名单的数组，红名单格式为：
        {"number": "120",
        "owerName": "lili",
        "longitude": "100.08",
        "latitude": "20.09",
        "owerAddress": "hubeiwuhan",
        "timeLen": 0}
        :return: 如果成功则返回”OK”，否则返回错误码
        """
        try:
            result = self.instance.AddToRedList(str(redLists))
            return result
        except BaseException as e:
            print(u"添加红名单出现异常：{0}".format(e))

    def QueryRedList(self, beginTime, endTime, status=1):
        """
        @description: 查询红名单
        :param beginTime: 开始时间
        :param endTime: 结束时间
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.QueryRedList(
                str(beginTime), str(endTime), str(status))
            return result
        except BaseException as e:
            print(u"查询红名单出现异常：{0}".format(e))

    def DeleteFromRedList(self, number):
        """
        @description: 删除红名单
        :number: 红名单的号码
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            number = json.dumps(number)
            print("number is:{0}".format(number))
            result = self.instance.DeleteFromRedList(number)
            return result
        except BaseException as e:
            print(u"删除红名单出现异常：{0}".format(e))

    def SetIdle(self):
        """
        @description: 置闲
        :return: 置闲成功以OK，否则返回错误码
        """
        try:
            result = self.instance.SetIdle()
            return result
        except BaseException as e:
            print(u"置闲出现异常：{0}".format(e))

    def SetBusy(self):
        """
        @description: 置忙
        :return: 置忙成功以OK，否则返回错误码
        """
        try:
            result = self.instance.SetBusy()
            return result
        except BaseException as e:
            print(u"置忙出现异常：{0}".format(e))

    def select_pick_up(self, number, quequeName):
        """
        @description: 选接
        :number: 目标号码
        :number: 队列名
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            print("number is:{0}".format(number))
            result = self.instance.SelectPickup(str(number), str(quequeName))
            return result
        except BaseException as e:
            print(u"选接出现异常：{0}".format(e))

    def create_meeting(self):
        """
        @description: 创建会议
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingCreate()
            return result
        except BaseException as e:
            print(u"创建会议出现异常：{0}".format(e))

    def invite_to_meeting(self, numberList):
        """
        @description: 邀请加入会议
        :numberList,邀请加入会议的号码的数组，如[1000,1001,1002]
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingInvite(str(numberList))
            return result
        except BaseException as e:
            print(u"邀请加入会议出现异常：{0}".format(e))

    def accept_to_meeting(self):
        """
        @description: 接收加入会议
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingAccept()
            return result
        except BaseException as e:
            print(u"接收加入会议出现异常：{0}".format(e))

    def reject_to_meeting(self):
        """
        @description: 拒绝加入会议
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingReject()
            return result
        except BaseException as e:
            print(u"拒绝加入会议出现异常：{0}".format(e))

    def remove_meeting_numbers(self, numberList):
        """
        @description: 移除会议成员
        :numberList,移除会议的号码的数组，如[1000,1001,1002]
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingRemove(str(numberList))
            return result
        except BaseException as e:
            print(u"移除会议成员出现异常：{0}".format(e))

    def meeting_set_own(self, number):
        """
        @description: 转让会议主席
        :number,目标号码，如1000
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingSetOwn(str(number))
            return result
        except BaseException as e:
            print(u"转让会议主席出现异常：{0}".format(e))

    def leave_meeting(self):
        """
        @description: 离开会议
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingLeave()
            return result
        except BaseException as e:
            print(u"离开会议出现异常：{0}".format(e))

    def meeting_finish(self):
        """
        @description: 离开会议
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        """
        try:
            result = self.instance.MeetingFinish()
            return result
        except BaseException as e:
            print(u"结束会议出现异常：{0}".format(e))

    def snatch_pick_up(self, number, quequeName):
        """
        @description: 代接坐席
        :number: 目标号码
        :quequeName: 队列名
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        ps:高级坐席代接后，被代接坐席会受到系统主动返回的onBeSnatchPickUp的事件。
        若新来电为内线电话，则不能被代接。
        """
        try:
            result = self.instance.SnatchPickup(str(number), str(quequeName))
            return result
        except BaseException as e:
            print(u"代接坐席出现异常：{0}".format(e))

    def transfer_to_ivr(self, bReleaseTrans, accessCode):
        """
        @description: 转到IVR
        :bReleaseTrans: 布尔型，True：挂起转，False：释放转
        :accessCode: IVR流程接入码
        :return: 查询成功以OK@后面跟随json格式返回，否则返回错误码
        ps:高级坐席代接后，被代接坐席会受到系统主动返回的onBeSnatchPickUp的事件。
        若新来电为内线电话，则不能被代接。
        """
        try:
            result = self.instance.TransferToIVR(
                bool(bReleaseTrans), str(accessCode))
            return result
        except BaseException as e:
            print(u"转到IVR出现异常：{0}".format(e))

    def get_event_info(self, event_type=None, event_info_path=None):
        """
        @description: 读文件中的内容
        :param event_info_path: 文件的路径
        :param event_type: 事件类型，为空时代表类型
        :return: 事件列表，列表中的内容是事件相关的字典, eg:[{'OnStatusChange': {'status': 'IDLE', 'number': '1001'}}]
        """
        if not event_info_path:
            event_info_path = self.event_info_path
        print(event_info_path)
        event_file = open(event_info_path, 'r')
        event_info = event_file.readline()
        if not event_info:
            raise ValueError
        else:
            event_info_list = []
            while event_info:
                event_info = event_info.strip('\n')
                event_info = eval(event_info)
                if not event_type:
                    event_info_list.append(event_info)
                elif event_type in event_info.keys():
                    event_info_list.append(event_info)
                event_info = event_file.readline()
            event_file.close()
            print(event_info_list)
            return event_info_list

    def clear_event(self, event_info_path=None):
        """
        @description: 清空文件中的内容
        :param event_info_path: 文件的路径
        :return:
        """
        if not event_info_path:
            event_info_path = self.event_info_path
        print(event_info_path)
        open(event_info_path, "w").close()
        print("event_info信息清除成功")

    @staticmethod
    def check_time(input_time, compare_time='', time_delta=2):
        """
        获取获取当前时间：2021-02-03 15:59:52
        :param time_delta: 相对时间差值，输入方式为：2,
        :return_time 返回时间2021-02-03 15:59:52
        """
        time_format = '%Y-%m-%d %H:%M:%S'
        if not compare_time:
            compare_time = datetime.datetime.now()
        else:
            compare_time = datetime.datetime.strptime(
                compare_time, time_format)
        current_time = compare_time.strftime(time_format)
        print(u"打印当前日期时间：{0}".format(current_time))
        forword_time = (
                compare_time +
                datetime.timedelta(
                    seconds=int(time_delta))).strftime(time_format)
        back_time = (
                compare_time +
                datetime.timedelta(
                    seconds=-
                    int(time_delta))).strftime(time_format)
        if back_time < input_time < forword_time:
            print(u"时间在误差范围内，检查通过")
            return current_time
        else:
            print(u"时间不在误差范围内，检查不通过")
            return False


if __name__ == '__main__':
    ippbx = BasicCall()
    RobotRemoteServer(ippbx, *sys.argv[1:])
    ippbx.int_instance()
    # res=ippbx.user_login("10.10.5.20","jiaotou","2004","666666","11.114.0.12",18083)
    # print("login res:{0}".format(res))
    # ippbx.user_logout()
    # ippbx.check_time("2021-02-03 16:24:52", "2021-02-03 16:24:52")

    # time.sleep(10)

    # print("login res:{0}".format(res))
    # res=ippbx.call_out("18810639520")
    # print "callcallcallcallcall"
    # time.sleep(5)
    # res=ippbx.hang_up()
    # pick_up_thread.start()
    # time.sleep(20)
