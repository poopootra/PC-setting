import datetime

import numpy as np
import pandas as pd
import PySimpleGUI as sg
import win32com.client as win32
from win32com.client import Dispatch

# 曜日フィルター
weekdays_map = {
    "Mon": 0,
    "Tue": 1,
    "Wed": 2,
    "Thu": 3,
    "Fri": 4,
    "Sat": 5,
    "Sun": 6,
}


class UserInputHandler:
    def __init__(self):
        self.user_input = None
        self.format_date = "%Y-%m-%d"
        self.format_time = "%H:%M"

    # UIでユーザー入力を受け取る
    def get_user_input(self):
        layout = [
            [sg.Text("予定のタイトル:"), sg.InputText(key="title")],
            [sg.Text("メールアドレス (カンマ区切り):"), sg.InputText(key="emails")],
            [
                sg.Text("開始期間:"),
                sg.InputText(key="start_date"),
                sg.CalendarButton("選択", target="start_date", format=self.format_date),
            ],
            [
                sg.Text("終了期間:"),
                sg.InputText(key="end_date"),
                sg.CalendarButton("選択", target="end_date", format=self.format_date),
            ],
            [sg.Text("開始時刻 (HH:MM):"), sg.InputText("09:00", key="start_time")],
            [sg.Text("終了時刻 (HH:MM):"), sg.InputText("18:00", key="end_time")],
            [
                sg.Frame(
                    "対象曜日:",
                    [[sg.Checkbox(i, key=i) for i in weekdays_map.keys()]],
                )
            ],
            [
                sg.Text("最小単位:"),
                sg.Slider(
                    range=(30, 180),
                    default_value=60,
                    orientation="horizontal",
                    tick_interval=30,
                    resolution=30,
                    size=(20, 15),
                    key="min_duration",
                    tooltip="選択した最小単位",
                ),
            ],
            [
                sg.Text("単発か複数か:"),
                sg.Radio("単発", "type", default=True, key="single"),
                sg.Radio("複数", "type", key="multiple"),
            ],
            [sg.Submit("送信"), sg.Cancel("キャンセル")],
        ]

        window = sg.Window("予定送信プログラム", layout)
        while True:
            event, values = window.read()

            # ユーザーがウィンドウを閉じたりキャンセルしたりした場合
            if event == sg.WINDOW_CLOSED or event == "キャンセル":
                break

            # 必須項目がすべて入力されているかチェック
            required_fields = ["title", "emails", "start_date", "end_date"]
            if all(values[field] for field in required_fields) and any(
                values[day] for day in weekdays_map.keys()
            ):
                self.user_input = values
                break
            else:
                sg.popup("すべての項目を入力してください。", title="入力エラー")

        window.close()

    def put_weekdays_in_list(self):
        dict_temp = {i: self.user_input.pop(i) for i in weekdays_map.keys()}
        self.user_input["weekdays"] = [k for k, v in dict_temp.items() if v]

    def set_dtype(self):
        for i in ["start_date", "end_date"]:
            self.user_input[i] = datetime.datetime.strptime(
                self.user_input[i], self.format_date
            ).date()
        for i in ["start_time", "end_time"]:
            self.user_input[i] = datetime.datetime.strptime(
                self.user_input[i], self.format_time
            ).time()
        self.user_input["emails"] = self.user_input["emails"].split(",")


def check_if_available(
    mail: str,
    date_occurrence: datetime.date,
    start_time: datetime.time,
    end_time: datetime.time,
    duration: int,
) -> bool:
    outlook = Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    start_occurrence = int((start_time.hour * 60 - start_time.minute) / duration)
    necessary_occurrence = int(
        (
            end_time.hour * 60
            + end_time.minute
            - start_time.hour * 60
            - start_time.minute
        )
        / duration
    )

    try:
        # 共有カレンダーの取得
        recipient = namespace.CreateRecipient(mail)
        str_free_busy = recipient.FreeBusy(
            datetime.datetime.combine(date_occurrence, start_time), duration, True
        )
        return list(
            map(
                int,
                str_free_busy[
                    start_occurrence : start_occurrence + necessary_occurrence
                ],
            )
        )

    except Exception as e:
        print(f"Error accessing shared calendar for {mail}: {e}")
        # エラーが発生した場合はFalseを返す
        return True


class ScheduleHandler:
    def __init__(self, start_time, end_time, duration):
        self.df = None
        self.df_start_end = None
        self.start_time = start_time
        self.end_time = end_time
        self.duration = duration

    def convert_to_df(self, mails, start_date, end_date, weekdays):
        # 曜日フィルターの適用
        weekday_nums = [weekdays_map[day] for day in weekdays]
        date_range = pd.date_range(start=start_date, end=end_date)
        filtered_dates = [
            date.date() for date in date_range if date.weekday() in weekday_nums
        ]

        # 空のDataFrameを作成
        self.df = pd.DataFrame(
            index=filtered_dates,
            columns=mails,
        )
        self.df.index.name = "Date"

    def set_availability(self):
        for idx, row in self.df.iterrows():
            for i in range(row.shape[0]):
                row.iloc[i] = check_if_available(
                    row.index[i],
                    idx,
                    self.start_time,
                    self.end_time,
                    self.duration,
                )
        self.df = self.df.explode(list(self.df.columns))

    def set_time_in_idx(self):
        time_delta = datetime.timedelta(minutes=self.duration)
        time_range = []
        current_time = datetime.datetime.combine(datetime.date.today(), self.start_time)
        while current_time.time() < self.end_time:
            time_range.append(current_time.time())
            current_time += time_delta
        time_range = time_range * self.df.index.nunique()
        self.df["Time"] = time_range
        self.df.set_index("Time", append=True, inplace=True)
        self.df["All available"] = self.df.apply(lambda x: all(x.isin([0, 1])), axis=1)

    def detect_windows_to_send_blocks(self, n):
        self.df["Block to send"] = self.df.groupby("Date")["All available"].transform(
            lambda x: self.detect_consecutive_true(x, n)
        )

    def detect_consecutive_true(self, series, n):
        result = pd.Series(False, index=series.index)
        consecutive_count = 0
        for i, value in enumerate(series):
            if value:
                consecutive_count += 1
                if consecutive_count >= n:
                    result.iloc[i - n + 1 : i + 1] = True
            else:
                consecutive_count = 0
        return result

    def set_start_and_end(self):
        df_temp = self.df[
            self.df.groupby("Date")["Block to send"].transform(lambda x: x != x.shift())
        ].copy()
        df_temp = df_temp[
            (df_temp.index.get_level_values(1) != self.start_time)
            | (df_temp["All available"])
        ]
        df_date = df_temp.groupby("Date").size()
        idx = df_date[df_date.mod(2) != 0].index
        df_new = pd.DataFrame(index=idx)
        df_new["Time"] = self.end_time
        df_new["All available"] = False
        df_new.set_index("Time", append=True, inplace=True)
        df_temp = pd.concat([df_temp, df_new]).sort_index()
        datetime_index = df_temp.index.map(
            lambda x: datetime.datetime.combine(x[0], x[1])
        )
        even_index_values = datetime_index[::2]
        odd_index_values = datetime_index[1::2]
        self.df_start_end = pd.DataFrame(
            {"start": even_index_values.values, "end": odd_index_values.values}
        )


class OutlookHandler:
    def send_block(self, df, subject, attendees):
        outlook = win32.Dispatch("Outlook.Application")
        for _, row in df.iterrows():
            appointment = outlook.CreateItem(1)
            appointment.Start = row["start"].strftime("%Y-%m-%d %H:%M")
            appointment.Subject = subject
            appointment.Duration = (row["end"] - row["start"]).seconds // 60
            appointment.Location = "Block"
            appointment.RequiredAttendees = ";".join(attendees)
            appointment.ResponseRequested = False
            appointment.MeetingStatus = 1
            appointment.Save()
            appointment.Send()


if __name__ == "__main__":
    user_input = UserInputHandler()
    data = user_input.get_user_input()
    user_input.put_weekdays_in_list()
    user_input.set_dtype()
    schedule_handler = ScheduleHandler(
        user_input.user_input["start_time"],
        user_input.user_input["end_time"],
        30,
    )
    schedule_handler.convert_to_df(
        user_input.user_input["emails"],
        user_input.user_input["start_date"],
        user_input.user_input["end_date"],
        user_input.user_input["weekdays"],
    )
    schedule_handler.set_availability()
    schedule_handler.set_time_in_idx()
    schedule_handler.detect_windows_to_send_blocks(
        int(user_input.user_input["min_duration"] / 30)
    )
    schedule_handler.set_start_and_end()
    outlook_handler = OutlookHandler()
    outlook_handler.send_block(
        schedule_handler.df_start_end,
        user_input.user_input["title"],
        user_input.user_input["emails"],
    )
    print("送信完了しました")
