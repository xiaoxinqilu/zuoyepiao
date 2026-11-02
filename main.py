# -*- coding: utf-8 -*-
"""
作业票查询工具 V3（单文件 · Kivy）
====================================
"""
import os
import csv
import hashlib
import threading
import datetime

import requests
import urllib3
urllib3.disable_warnings()

HEADERS_CN = ['作业名称', '作业类型', '属地单位', '截止时间', '状态']
HOST = 'www.zyxk.qlsh:6030'
TENANT_ID = 2000000002403

# =====================================================================
# 一、加密 & 导出
# =====================================================================
def rsa_encrypt(text, modulus_hex):
    n = int(modulus_hex, 16)
    e = int('010001', 16)
    m = int.from_bytes(text.encode('utf-8'), 'big')
    return pow(m, e, n).to_bytes(128, 'big').hex()

def calc_password(password, modulus):
    sha = hashlib.sha256(password.encode('utf-8')).hexdigest().upper()
    return rsa_encrypt(sha, modulus), rsa_encrypt(password, modulus)

def sort_rows(records):
    def key(r):
        t = (r.get('piao_time') or '').strip()
        return (t == '', t)
    return sorted(records, key=key)

def _row_values(r):
    status = '正在作业' if r.get('is_pause', 0) == 0 else '已暂停'
    return [r.get('workname', ''), r.get('worktype_name', ''),
            r.get('territorialunitname', ''), r.get('piao_time', ''), status]

def export_csv(records, path):
    records = sort_rows(records)
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS_CN)
        for r in records:
            writer.writerow(_row_values(r))
    return path

def export_xls(records, path):
    records = sort_rows(records)
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = '作业票'
    ws.append(HEADERS_CN)
    header_fill = PatternFill('solid', fgColor='2166A6')
    header_font = Font(color='FFFFFF', bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for r in records:
        ws.append(_row_values(r))
    widths = [40, 14, 20, 20, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = 'A2'
    wb.save(path)
    return path

# =====================================================================
# 二、网络客户端（含重试）
# =====================================================================
class TicketClient:
    def __init__(self, username, password, host=HOST, tenantid=TENANT_ID, timeout=15):
        self.username = username
        self.password = password
        self.host = host
        self.tenantid = tenantid
        self.timeout = timeout
        self.session = requests.Session()
        self.session.verify = False
        try:
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
            retry = Retry(total=2, backoff_factor=0.4, status_forcelist=[500, 502, 503, 504])
            self.session.mount('https://', HTTPAdapter(max_retries=retry))
        except Exception:
            pass

    def _base_headers(self):
        return {'Host': self.host, 'Connection': 'Keep-Alive',
                'Accept-Encoding': 'gzip', 'User-Agent': 'okhttp/3.8.0'}

    def login(self):
        url_get = f'https://{self.host}/m/passport/login/getEncryptType.json'
        resp = self.session.get(url_get, headers=self._base_headers(), timeout=self.timeout)
        resp.raise_for_status()
        data = resp.json()
        if not data.get('data'):
            raise RuntimeError(f"获取加密信息失败：{data.get('msg', '未知')}")
        login_stoken = data['data']['loginStoken']
        modulus = data['data']['pubKeyVO']['modulus']
        pw_value, vp_value = calc_password(self.password, modulus)

        url_login = f'https://{self.host}/m/passport/login/login.json'
        headers1 = {'Accept': 'application/json', 'Content-Type': 'application/json',
                    'Host': self.host, 'Connection': 'Keep-Alive',
                    'Accept-Encoding': 'gzip', 'User-Agent': 'okhttp/3.8.0'}
        payload = {
            "appVersion": "01.23.0111", "equipmentCode": "867632029715366",
            "ip": "172.16.1.15", 'isAD': True, "loginStoken": login_stoken,
            "password": pw_value, "tenantid": self.tenantid,
            "username": self.username, "verifyPwd": vp_value,
        }
        result = self.session.post(url_login, json=payload, headers=headers1, timeout=self.timeout).json()
        if not result.get('data'):
            raise RuntimeError(f"登录失败：{result.get('msg', '账号或密码错误')}")
        return result['data']['st']

    def fetch_tickets(self, rows=500):
        st = self.login()
        url = f'https://{self.host}/m/hse_m/HSE_WORKTASK_INJOB_M/listQuery.json?queryType=2&level=2&page=1&rows={rows}&extWhere='
        headers2 = {
            'st': st, 'Accept': 'application/json', 'tid': str(self.tenantid),
            'Content-Type': 'application/json;charset=UTF-8',
            'user-agent': '2211133C(Android/9) (com.hayden.hap.fv/1.0.2) Weex/0.16.0 900x1600',
            'Host': self.host, 'Connection': 'Keep-Alive', 'Accept-Encoding': 'gzip',
        }
        resp = self.session.get(url, headers=headers2, timeout=self.timeout)
        resp.raise_for_status()
        data = resp.json().get('data') or []
        records = []
        for item in data:
            try:
                vo = item['children']['HSE_WORKTICKET_INJOB_M'][0]['vo']
            except (KeyError, IndexError, TypeError):
                continue
            records.append({
                'workname': vo.get('workname', ''),
                'worktype_name': vo.get('worktype_name', ''),
                'territorialunitname': vo.get('territorialunitname', ''),
                'piao_time': vo.get('valid_closetime', '') or '',
                'is_pause': vo.get('is_pause', 0),
            })
        return sort_rows(records)

# =====================================================================
# 三、Kivy UI
# =====================================================================
from kivy.app import App
from kivy.core.window import Window
from kivy.core.text import LabelBase
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.graphics import Color, Rectangle, Line
from kivy.utils import platform

def _register_cn_font():
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base, 'NotoSansSC-Regular.ttf'),
        os.path.join(base, 'app', 'NotoSansSC-Regular.ttf'),
        os.path.join(base, 'assets', 'NotoSansSC-Regular.ttf'),
    ]
    if os.name == 'nt':
        candidates += [r'C:\Windows\Fonts\msyh.ttc', r'C:\Windows\Fonts\simhei.ttf']
    if platform == 'android':
        candidates += ['/system/fonts/NotoSansCJK-Regular.ttc', '/system/fonts/DroidSansFallback.ttf']
    for p in candidates:
        if os.path.exists(p):
            try:
                LabelBase.register(name='CNFont', fn_regular=p)
                return 'CNFont'
            except Exception:
                continue
    return None

CN_FONT = _register_cn_font()

if platform == 'android':
    from android.permissions import request_permissions, Permission
else:
    request_permissions = None

def get_save_dir():
    if platform == 'android':
        try:
            from jnius import autoclass
            Env = autoclass('android.os.Environment')
            p = Env.getExternalStoragePublicDirectory(Env.DIRECTORY_DOWNLOADS).getAbsolutePath()
            if os.path.isdir(p) and os.access(p, os.W_OK):
                return p
        except Exception:
            pass
        try:
            return App.get_running_app().user_data_dir
        except Exception:
            return os.path.expanduser('~')
    return os.path.dirname(os.path.abspath(__file__))

def _history_path():
    if platform == 'android':
        try:
            from jnius import autoclass
            A = autoclass('org.kivy.android.PythonActivity')
            return os.path.join(A.mActivity.getFilesDir().getAbsolutePath(), 'user_history.json')
        except Exception:
            pass
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'user_history.json')

def load_history():
    import json
    try:
        with open(_history_path(), encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, list):
            return [str(x) for x in data if str(x).strip()]
    except Exception:
        pass
    return []

def save_history(users):
    import json
    seen, uniq = set(), []
    for u in users:
        if u and u.strip() and u not in seen:
            seen.add(u)
            uniq.append(u)
    uniq = uniq[:20]
    try:
        with open(_history_path(), 'w', encoding='utf-8') as f:
            json.dump(uniq, f, ensure_ascii=False)
    except Exception:
        pass
    return uniq

def add_history(username):
    users = load_history()
    users = [username] + [u for u in users if u != username]
    return save_history(users)

def is_urgent(piao_time, now=None):
    t = (piao_time or '').strip().replace('/', '-')
    if not t:
        return False
    dt = None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            dt = datetime.datetime.strptime(t, fmt)
            break
        except ValueError:
            continue
    if dt is None:
        return False
    now = now or datetime.datetime.now()
    tomorrow = now.date() + datetime.timedelta(days=1)
    deadline = datetime.datetime.combine(tomorrow, datetime.time(15, 0, 0))
    return dt <= deadline

class WrapLabel(Label):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.halign = 'center'
        self.valign = 'middle'
        self.bind(width=self._rewrap)

    def _rewrap(self, *a):
        self.text_size = (max(self.width - dp(10), dp(10)), None)

    def need_height(self):
        return max(dp(30), self.texture_size[1] + dp(10))

class CNTextInput(TextInput):
    def keyboard_on_key_down(self, window, keycode, text, modifiers):
        key = keycode[1] if keycode else ''
        if key in ('backspace', 'delete'):
            if self.selection:
                self.delete_selection()
                return True
            ci = self.cursor_index()
            if key == 'backspace' and ci > 0:
                self.text = self.text[:ci - 1] + self.text[ci:]
                self.cursor = self.get_cursor_from_index(ci - 1)
                return True
            if key == 'delete' and ci < len(self.text):
                self.text = self.text[:ci] + self.text[ci + 1:]
                self.cursor = self.get_cursor_from_index(ci)
                return True
            return True
        return super().keyboard_on_key_down(window, keycode, text, modifiers)

class ZhuyepiaoApp(App):
    title = '作业票查询'

    def build(self):
        Window.softinput_mode = 'below_target'
        Window.clearcolor = (1, 1, 1, 1)
        self._busy = False

        root = BoxLayout(orientation='vertical', padding=dp(8), spacing=dp(6))

        user_row = BoxLayout(size_hint_y=None, height=dp(42), spacing=dp(6))
        self.history = load_history()
        self.txt_user = CNTextInput(
            text=(self.history[0] if self.history else ''), multiline=False, hint_text='账号',
            font_name=CN_FONT, background_color=(0.96, 0.96, 0.96, 1), foreground_color=(0.1, 0.1, 0.1, 1))
        self.txt_user.bind(text=self._on_user_typing)
        self.txt_pwd = CNTextInput(
            text='123456', multiline=False, password=True, hint_text='密码',
            font_name=CN_FONT, background_color=(0.96, 0.96, 0.96, 1), foreground_color=(0.1, 0.1, 0.1, 1))
        user_row.add_widget(self.txt_user)

        btn_hist = Button(text='历史', font_name=CN_FONT, size_hint_x=None, width=dp(52), background_color=(0.75, 0.78, 0.85, 1))
        btn_hist.bind(on_release=self._show_history)
        user_row.add_widget(btn_hist)
        user_row.add_widget(self.txt_pwd)
        root.add_widget(user_row)

        btn_row = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(6))
        self.btn_query = Button(text='查询作业票', font_name=CN_FONT, background_color=(0.13, 0.55, 0.9, 1))
        self.btn_query.bind(on_press=self.on_query)
        self.btn_csv = Button(text='导出CSV', font_name=CN_FONT, background_color=(0.2, 0.65, 0.35, 1))
        self.btn_csv.bind(on_press=self.on_export_csv)
        self.btn_xls = Button(text='导出XLS', font_name=CN_FONT, background_color=(0.85, 0.55, 0.1, 1))
        self.btn_xls.bind(on_press=self.on_export_xls)
        for b in (self.btn_query, self.btn_csv, self.btn_xls):
            btn_row.add_widget(b)
        root.add_widget(btn_row)

        filter_row = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(6))
        filter_row.add_widget(Label(text='属地单位筛选：', font_name=CN_FONT, size_hint_x=None, width=dp(110), color=(0.15, 0.15, 0.15, 1), font_size=dp(14)))
        self.btn_filter = Button(text='全部 ▼', font_name=CN_FONT, background_color=(0.93, 0.95, 1, 1), color=(0.1, 0.1, 0.1, 1))
        self.btn_filter.bind(on_release=self._show_unit_filter)
        filter_row.add_widget(self.btn_filter)
        root.add_widget(filter_row)

        self.lbl_status = Label(font_name=CN_FONT, text='请先查询', size_hint_y=None, height=dp(26), color=(0.35, 0.35, 0.35, 1), font_size=dp(13))
        root.add_widget(self.lbl_status)

        self.table = GridLayout(cols=5, size_hint_y=None, spacing=0, row_default_height=dp(30), row_force_default=False)
        self.table.bind(minimum_height=self.table.setter('height'))
        self.scroll = ScrollView(bar_color=(0.6, 0.6, 0.6, 1))
        self.scroll.add_widget(self.table)
        root.add_widget(self.scroll)

        self.rows = []

        if platform == 'android' and request_permissions:
            request_permissions([Permission.INTERNET, Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
        return root

    def _show_history(self, *args):
        from kivy.uix.popup import Popup
        hist = getattr(self, 'history', [])
        if not hist:
            self._set_status('暂无历史账号，先查询一次即可记录')
            return
        box = BoxLayout(orientation='vertical', spacing=dp(4), padding=dp(6))
        sv = ScrollView()
        inner = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(4))
        inner.bind(minimum_height=inner.setter('height'))
        popup = Popup(title='选择历史账号', content=box, size_hint=(0.85, 0.7), title_font=CN_FONT)
        for u in hist:
            b = Button(text=u, font_name=CN_FONT, size_hint_y=None, height=dp(44), background_color=(0.9, 0.93, 1, 1), color=(0.1, 0.1, 0.1, 1))
            b.bind(on_release=lambda inst, name=u: self._pick_user(name, popup))
            inner.add_widget(b)
        sv.add_widget(inner)
        box.add_widget(sv)
        close = Button(text='关闭', font_name=CN_FONT, size_hint_y=None, height=dp(42), background_color=(0.85, 0.85, 0.85, 1))
        close.bind(on_release=lambda *a: popup.dismiss())
        box.add_widget(close)
        popup.open()

    def _pick_user(self, name, popup):
        self.txt_user.text = name
        popup.dismiss()

    def _on_user_typing(self, inst, value):
        v = (value or '').strip()
        hist = getattr(self, 'history', [])
        if not v:
            inst.hint_text = (f"账号（历史：{'/'.join(hist[:3])}）" if hist else '账号')
            return
        matches = [u for u in hist if u != v and u.startswith(v)]
        inst.hint_text = ('联想：' + matches[0]) if matches else '账号'

    def _set_status(self, text, *args):
        self.lbl_status.text = text

    def _cell(self, text, bg=None, bold=True, font_size_dp=13, red=False):
        if red:
            fg = (0.85, 0.05, 0.05, 1)
        elif bg and bg[0] < 0.5:
            fg = (1, 1, 1, 1)
        else:
            fg = (0, 0, 0, 1)
        lab = WrapLabel(text=str(text), bold=bold, font_name=CN_FONT, font_size=dp(font_size_dp), color=fg, size_hint_y=None)
        if bg:
            with lab.canvas.before:
                Color(*bg)
                lab._bg_rect = Rectangle(pos=lab.pos, size=lab.size)
        with lab.canvas.after:
            Color(0, 0, 0, 1)
            lab._border = Line(rectangle=(lab.x, lab.y, lab.width, lab.height), width=1.2)

        def _update(inst, *_):
            if bg:
                inst._bg_rect.pos = inst.pos
                inst._bg_rect.size = inst.size
            inst._border.rectangle = (inst.x, inst.y, inst.width, inst.height)
        lab.bind(pos=_update, size=_update)
        return lab

    def _add_row(self, values, bg, red=False, bold=True, font_size_dp=13):
        cells = [self._cell(v, bg=bg, red=red, bold=bold, font_size_dp=font_size_dp) for v in values]
        for c in cells:
            self.table.add_widget(c)
        return cells

    def _unify_heights(self, all_cells):
        for cells in all_cells:
            try:
                hmax = max([c.need_height() for c in cells] + [dp(30)])
                for c in cells:
                    c.height = hmax
            except Exception:
                pass

    def _fill_table(self, rows):
        self.table.clear_widgets()
        self.table.cols = 5
        all_cells = []
        all_cells.append(self._add_row(HEADERS_CN, bg=(0.13, 0.4, 0.65, 1), bold=True, font_size_dp=14))
        for idx, r in enumerate(rows):
            status = '正在作业' if r['is_pause'] == 0 else '已暂停'
            stripe = (0.93, 0.96, 1.0, 1) if idx % 2 == 0 else (1, 1, 1, 1)
            urgent = is_urgent(r.get('piao_time'))
            values = (r['workname'], r['worktype_name'], r['territorialunitname'], r['piao_time'], status)
            all_cells.append(self._add_row(values, bg=stripe, red=urgent))
        Clock.schedule_once(lambda dt: self._unify_heights(all_cells), 0)
        Clock.schedule_once(lambda dt: self._unify_heights(all_cells), 0.1)
        try:
            self.scroll.scroll_y = 1
        except Exception:
            pass

    def _show_unit_filter(self, *args):
        from kivy.uix.popup import Popup
        rows = getattr(self, 'rows', None)
        if not rows:
            self._set_status('请先查询，再筛选属地单位')
            return
        units = sorted({r['territorialunitname'] for r in rows if r['territorialunitname']})
        box = BoxLayout(orientation='vertical', spacing=dp(4), padding=dp(6))
        sv = ScrollView()
        inner = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(4))
        inner.bind(minimum_height=inner.setter('height'))
        popup = Popup(title='选择属地单位', content=box, size_hint=(0.85, 0.7), title_font=CN_FONT)

        def _mk(unit):
            b = Button(text=unit, font_name=CN_FONT, size_hint_y=None, height=dp(44), background_color=(0.9, 0.93, 1, 1), color=(0.1, 0.1, 0.1, 1))
            b.bind(on_release=lambda inst: self._pick_unit(unit, popup))
            return b

        inner.add_widget(_mk('全部'))
        for u in units:
            inner.add_widget(_mk(u))
        sv.add_widget(inner)
        box.add_widget(sv)
        close = Button(text='关闭', font_name=CN_FONT, size_hint_y=None, height=dp(42), background_color=(0.85, 0.85, 0.85, 1))
        close.bind(on_release=lambda *a: popup.dismiss())
        box.add_widget(close)
        popup.open()

    def _pick_unit(self, unit, popup):
        self.btn_filter.text = (unit + ' ▼') if unit != '全部' else '全部 ▼'
        self._apply_unit(unit)
        popup.dismiss()

    def _apply_unit(self, unit):
        rows = getattr(self, 'rows', None)
        if not rows:
            return
        shown = rows if unit == '全部' else [r for r in rows if r['territorialunitname'] == unit]
        self._fill_table(shown)
        urgent_n = sum(1 for r in shown if is_urgent(r.get('piao_time')))
        self._set_status(f"筛选【{unit}】：{len(shown)} / {len(rows)} 条｜红色 {urgent_n} 条临期")

    def on_query(self, *args):
        if self._busy:
            return
        user = self.txt_user.text.strip()
        pwd = self.txt_pwd.text.strip()
        if not user or not pwd:
            self._set_status('请填写账号和密码')
            return
        self._busy = True
        self.btn_query.disabled = True
        self._set_status('正在连接内网查询…')
        threading.Thread(target=self._do_query, args=(user, pwd), daemon=True).start()

    def _do_query(self, user, pwd):
        try:
            client = TicketClient(user, pwd)
            rows = client.fetch_tickets()
            Clock.schedule_once(lambda dt: self._after_query(user, rows), 0)
        except Exception as e:
            err = str(e) or e.__class__.__name__
            Clock.schedule_once(lambda dt: self._query_fail(err), 0)

    def _after_query(self, user, rows):
        self._busy = False
        self.btn_query.disabled = False
        self.rows = rows
        try:
            if user:
                self.history = add_history(user)
        except Exception:
            pass
        self._fill_table(rows)
        self.btn_filter.text = '全部 ▼'
        urgent_n = sum(1 for r in rows if is_urgent(r.get('piao_time')))
        self._set_status(f"查询成功，共 {len(rows)} 条（时间升序）｜红色 {urgent_n} 条临期")

    def _query_fail(self, err):
        self._busy = False
        self.btn_query.disabled = False
        short = err if len(err) <= 60 else err[:60] + '…'
        self._set_status(f"查询失败：{short}（请确认已连内网 VPN/WiFi）")

    def _do_export(self, func, ext):
        if not getattr(self, 'rows', None):
            self._set_status('请先查询再导出')
            return
        filename = '作业票封闭.' + ext
        try:
            path = os.path.join(get_save_dir(), filename)
            func(self.rows, path)
            self._set_status('已导出：' + path)
            return
        except Exception:
            pass
        try:
            path = os.path.join(self.user_data_dir, filename)
            func(self.rows, path)
            self._set_status('已导出到应用目录：' + path)
        except Exception as e2:
            self._set_status('导出失败：' + str(e2))

    def on_export_csv(self, *args):
        self._do_export(export_csv, 'csv')

    def on_export_xls(self, *args):
        self._do_export(export_xls, 'xlsx')

if __name__ == '__main__':
    ZhuyepiaoApp().run()