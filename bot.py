import re
import json
import logging
import traceback

import telebot
from telebot import types

from code import Code
from office_user import OfficeUser

bot = telebot.TeleBot(
    token='1768910084:AAGajN---_xO5LQnYM9Bs-d1fNxAqio5RDs',
    parse_mode='HTML'
)

config = json.load(open('config.json'))

user_dict = {
    # 'user_id': {
    #     'selected_sub': {},
    #     'selected_domain': '',
    #     'username': '',
    #     'code': ''
    # }
}

OU = OfficeUser(
    client_id=config['aad']['clientId'],
    tenant_id=config['aad']['tenantId'],
    client_secret=config['aad']['clientSecret']
)
C = Code()


def start(m):
    if m.from_user.id == config['bot']['admin']:
        bot.send_message(
            text='欢迎使用 <b>Office User Bot</b>\n\n'
                 '可用的命令有：\n'
                 '/create 创建 Office 账号\n'
                 '/gen 10 生成十个激活码\n'
                 '/about 关于 Bot',
            chat_id=m.from_user.id
        )

    else:
        bot.send_message(
            text='欢迎使用 <b>Office User Bot</b>\n\n'
                 '可用的命令有：\n'
                 '/create 创建 Office 账号\n'
                 '/about 关于 Bot',
            chat_id=m.from_user.id
        )


def gen(m):
    amount = int(str(m.text).strip().split('/gen')[1].strip())
    codes = C.gen(amount)
    bot.send_message(
        text='\n'.join(codes),
        chat_id=m.from_user.id
    )


def create(m):
    buttons = [types.KeyboardButton(
        text=sub['name']
    ) for sub in config['office']['subscriptions']]

    markup = types.ReplyKeyboardMarkup(row_width=1)
    markup.add(*buttons)
    msg = bot.send_message(
        text='欢迎创建 Office 账号\n\n请选择订阅：',
        chat_id=m.from_user.id,
        reply_markup=markup
    )
    bot.register_next_step_handler(msg, select_subscription)


def select_subscription(m):
    selected_sub = next(
        (sub for sub in config['office']['subscriptions'] if sub['name'] == m.text),
        None
    )
    if selected_sub is None:
        msg = bot.send_message(
            text='订阅不存在，请重新回复：',
            chat_id=m.from_user.id,
        )
        bot.register_next_step_handler(msg, select_subscription)
        return
    user_dict[m.from_user.id] = {}
    user_dict[m.from_user.id]['selected_sub'] = selected_sub

    markup = types.ReplyKeyboardRemove(selective=False)
    msg = bot.send_message(
        text='请回复想要的用户名：',
        chat_id=m.from_user.id,
        reply_markup=markup
    )
    bot.register_next_step_handler(msg, input_username)


def input_username(m):
    username = str(m.text).strip()
    if username in config['banned']['officeUsername'] or \
            not re.match(r'^[a-zA-Z0-9\-]+$', username):
        msg = bot.send_message(
            text='用户名含有特殊字符或在黑名单中，请重新回复：',
            chat_id=m.from_user.id,
        )
        bot.register_next_step_handler(msg, input_username)
        return
    user_dict[m.from_user.id]['username'] = username

    buttons = [types.KeyboardButton(
        text=d['display']
    ) for d in config['office']['domains']]
    markup = types.ReplyKeyboardMarkup(row_width=1)
    markup.add(*buttons)
    msg = bot.send_message(
        text='请选择账号后缀：',
        chat_id=m.from_user.id,
        reply_markup=markup
    )
    bot.register_next_step_handler(msg, select_domain)


def select_domain(m):
    selected_domain = next(
        (d for d in config['office']['domains'] if d['display'] == m.text),
        None
    )
    if selected_domain is None:
        msg = bot.send_message(
            text='后缀不存在，请重新回复：',
            chat_id=m.from_user.id,
        )
        bot.register_next_step_handler(msg, select_domain)
        return
    user_dict[m.from_user.id]['selected_domain'] = selected_domain

    markup = types.ReplyKeyboardRemove(selective=False)
    msg = bot.send_message(
        text='请回复激活码：',
        chat_id=m.from_user.id,
        reply_markup=markup
    )
    bot.register_next_step_handler(msg, input_code)


def input_code(m):
    code = str(m.text).strip()
    if not C.check(code):
        bot.send_message(
            text='激活码无效！',
            chat_id=m.from_user.id,
        )
        return
    # todo: lock code
    user_dict[m.from_user.id]['code'] = code

    selected_sub_name = user_dict[m.from_user.id]['selected_sub']['name']
    selected_domain_display = user_dict[m.from_user.id]['selected_domain']['display']
    username = user_dict[m.from_user.id]['username']

    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
        types.InlineKeyboardButton(text='取消', callback_data='cancel'),
        types.InlineKeyboardButton(text='确认', callback_data='create'),
    )
    bot.send_message(
        text=f'{selected_sub_name}\n'
             f'{username}@{selected_domain_display}\n\n'
             '激活码有效，确认创建账号吗？',
        chat_id=m.from_user.id,
        reply_markup=markup
    )


def notify_admin(call):
    if config['bot']['notify']:
        user_id = call.from_user.id

        selected_sub_name = user_dict[user_id]['selected_sub']['name']
        selected_domain_value = user_dict[user_id]['selected_domain']['value']
        username = user_dict[user_id]['username']
        code = user_dict[user_id]['code']
        tg_name = f'{call.from_user.first_name or ""} {call.from_user.last_name or ""}'.strip()

        bot.send_message(
            text=f'<a href="tg://user?id={user_id}">{tg_name}</a> 刚刚用激活码 {code} 创建了 '
                 f'{username}{selected_domain_value} ({selected_sub_name})',
            chat_id=config['bot']['admin']
        )


@bot.message_handler(content_types=['text'])
def handle_text(m):
    # noinspection PyBroadException
    try:
        if m.from_user.id in config['banned']['tgId']:
            return

        text = str(m.text).strip()

        bot.send_chat_action(
            chat_id=m.from_user.id,
            action='typing'
        )
        if text == '/create':
            create(m)

        elif text.startswith('/gen'):
            gen(m)

        else:
            start(m)

    except Exception:
        traceback.print_exc()


def create_account(call):
    user_id = call.from_user.id
    msg_id = call.message.message_id
    chat_id = call.from_user.id

    if user_dict.get(user_id) is None:
        return

    bot.edit_message_text(
        chat_id=chat_id,
        text='创建账号中，请稍等...',
        message_id=msg_id
    )

    try:
        account = OU.create_account(
            username=user_dict[user_id]['username'],
            domain=user_dict[user_id]['selected_domain']['value'],
            sku_id=user_dict[user_id]['selected_sub']['sku'],
            display_name=f'{call.from_user.first_name or ""} {call.from_user.last_name or ""}'.strip(),
        )
        C.del_code(user_dict[user_id]['code'])

        selected_sub_name = user_dict[user_id]['selected_sub']['name']
        bot.send_message(
            text='账号创建成功\n'
                 '===========\n\n'
                 f'订阅： {selected_sub_name}\n'
                 f'邮箱： <b>{account["email"]}</b>\n'
                 f'初始密码： <b>{account["password"]}</b>\n\n'
                 f'登录地址： https://office.com',
            chat_id=chat_id
        )

        notify_admin(call)
        del user_dict[user_id]

    except Exception as e:
        error = json.loads(str(e))
        if 'userPrincipalName already exists' in error['error']['message']:
            text = '用户名已存在，请换个用户名重新创建账号'

        else:
            text = '哎呀出错了'

        bot.send_message(
            text=text,
            chat_id=chat_id
        )


@bot.callback_query_handler(func=lambda call: True)
def handle_callback(call):
    if call.data == 'create':
        create_account(call)

    elif call.data == 'cancel':
        bot.edit_message_text(
            chat_id=call.from_user.id,
            text='已取消',
            message_id=call.message.message_id
        )


logger = telebot.logger
telebot.logger.setLevel(logging.DEBUG)

bot.polling(none_stop=True)
