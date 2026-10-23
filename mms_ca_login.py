"""Optional CA authentication through the installed company certificate client."""
from __future__ import annotations

import base64
import json
import os
from html.parser import HTMLParser
from urllib.parse import parse_qs, urlparse

SSO = "https://ldap.crc.com.cn"
MMS = "https://mms.crland.com.cn"


class _HiddenFields(HTMLParser):
    def __init__(self):
        super().__init__()
        self.fields = {}

    def handle_starttag(self, tag, attrs):
        values = dict(attrs)
        if tag == "input" and values.get("name") in {"code", "state"}:
            self.fields[values["name"]] = values.get("value", "")


def _json_response(response, stage):
    if not response.ok:
        raise RuntimeError(f"CA 登录失败：{stage} HTTP {response.status}")
    try:
        data = response.json()
    except Exception:
        raise RuntimeError(f"CA 登录失败：{stage} 未返回 JSON") from None
    if not isinstance(data, dict):
        raise RuntimeError(f"CA 登录失败：{stage} 响应格式异常")
    return data


def login_with_ca(context, *, timeout_seconds=60, connect=None, name=None, pin=None, identity_index=None):
    """Use the browser context's cookie jar; never guess a frontend token key."""
    name = (name or os.environ.get("MMS_CA_NAME", "")).strip()
    pin = pin or os.environ.get("MMS_CA_PIN", "")
    if not name or not pin:
        raise ValueError("请在 config.json 的 mmsApi 中填写 caName、caPin，或设置 MMS_CA_NAME 和 MMS_CA_PIN")
    try:
        identity_index = int(identity_index if identity_index is not None else os.environ.get("MMS_CA_POSITION", "0"))
        if identity_index < 0:
            raise ValueError
    except ValueError:
        raise ValueError("MMS_CA_POSITION 必须是非负整数") from None
    if connect is None:
        try:
            from websocket import create_connection
        except ImportError:
            raise RuntimeError("请安装 requirements-playwright.txt 中的 websocket-client") from None
        connect = create_connection
    try:
        socket = connect("wss://127.0.0.1:10022/", timeout=timeout_seconds)
    except Exception:
        raise RuntimeError("无法连接本机证书客户端 wss://127.0.0.1:10022/；请检查客户端、证书和运行账户") from None
    request = context.request
    timeout = timeout_seconds * 1000
    try:
        socket.send(json.dumps({"action": "enumerate_ukey_user"}))
        try:
            devices = json.loads(socket.recv())["data"]["usbkey"]
            device = devices[0]
            container = device["identity"][identity_index]["container"]
            key_index = device["keyindex"]
        except (ValueError, KeyError, IndexError, TypeError):
            raise RuntimeError("证书枚举失败：没有可用设备或证书序号不正确") from None
        challenge = _json_response(request.get(
            SSO + "/authn/authn/authFeishuRandom", timeout=timeout), "获取签名挑战")
        original = challenge.get("original_data")
        challenge_key = challenge.get("authnCerKey")
        if not isinstance(original, str) or not challenge_key:
            raise RuntimeError("CA 登录失败：签名挑战缺少必要字段")
        socket.send(json.dumps({
            "action": "sm9_sign", "pin": pin, "keyindex": key_index,
            "container": container, "message": base64.b64encode(original.encode()).decode(),
            "hashtype": "none",
        }))
        signed = json.loads(socket.recv()).get("data")
        if not signed or not isinstance(signed, str):
            raise RuntimeError("证书签名失败，请检查 PIN 和证书")
        auth = request.get(SSO + "/idp/authCenter/authenticate", params={
            "response_type": "code", "state": "1", "redirect_uri": MMS + "/login",
            "client_id": "CRLD_mms",
        }, timeout=timeout)
        if not auth.ok:
            raise RuntimeError(f"认证中心 HTTP {auth.status}")
        lck = parse_qs(urlparse(auth.url).query).get("lck", [""])[0]
        if not lck:
            raise RuntimeError("认证中心未返回 lck；需要核对当前认证流程")
        result = _json_response(request.post(SSO + "/authn/authn/authExecute?lang=zh_CN", data={
            "authChainCode": "3d56ea21707c4bbbbcb7c54e64b9a10f",
            "authModuleCode": "authAoLianCA", "entityId": "CRLD_mms",
            "requestType": "chain_type", "lck": lck,
            "authPara": {"authMode": "ca", "authnCerKey": challenge_key,
                         "original": original, "sid": name.upper(), "signed_data": signed,
                         "token": "", "loginName": name.upper()},
        }, timeout=timeout), "证书认证")
        token = result.get("loginToken")
        if not token:
            raise RuntimeError("CA 登录失败：认证中心未返回 loginToken")
        page = request.get(SSO + "/idp/authCenter/authnEngine",
                           params={"loginToken": token}, timeout=timeout)
        if not page.ok:
            raise RuntimeError(f"认证回调 HTTP {page.status}")
        parser = _HiddenFields()
        parser.feed(page.text())
        code = parser.fields.get("code")
        if not code:
            raise RuntimeError("认证回调缺少 code，需要核对认证页面")
        return _json_response(request.post(MMS + "/api/user-center/users/login",
                             data={"code": code, "loginType": "LDAP", "clientId": "CRLD_mms"}, timeout=timeout), "良域登录")
    finally:
        socket.close()
