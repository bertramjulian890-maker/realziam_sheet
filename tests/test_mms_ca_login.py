import json
import os
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from mms_ca_login import _HiddenFields, login_with_ca


def response(data=None, *, text="", url="", ok=True):
    return SimpleNamespace(ok=ok, status=200 if ok else 500, url=url,
                           json=lambda: data, text=lambda: text)


class CaLoginTest(unittest.TestCase):
    def test_authentication_uses_browser_cookie_session_and_closes_socket(self):
        socket = Mock()
        socket.recv.side_effect = [json.dumps({"data": {"usbkey": [
            {"keyindex": "key", "identity": [{"container": "cert"}]}]}}),
            json.dumps({"data": "signed-challenge"})]
        request = Mock()
        request.get.side_effect = [
            response({"original_data": "challenge", "authnCerKey": "challenge-key"}),
            response(url="https://ldap.crc.com.cn/auth?other=x&lck=abc%2B123&next=y"),
            response(text="<input value='code&amp;value' name='code'>"),
        ]
        request.post.side_effect = [response({"loginToken": "temporary-token"}),
                                    response({"data": {"user": "example"}})]
        with patch.dict(os.environ, {"MMS_CA_NAME": "example", "MMS_CA_PIN": "test-pin",
                                    "MMS_CA_POSITION": "0"}):
            result = login_with_ca(SimpleNamespace(request=request), connect=Mock(return_value=socket))
        self.assertIn("data", result)
        auth_payload = request.post.call_args_list[0].kwargs["data"]
        self.assertEqual(auth_payload["lck"], "abc+123")
        self.assertEqual(auth_payload["authPara"]["loginName"], "EXAMPLE")
        self.assertEqual(request.post.call_args_list[1].kwargs["data"]["code"], "code&value")
        self.assertEqual(request.post.call_args_list[1].kwargs["data"]["clientId"], "CRLD_mms")
        socket.close.assert_called_once()

    def test_missing_environment_stops_before_connecting(self):
        connect = Mock()
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaisesRegex(ValueError, "MMS_CA_NAME"):
                login_with_ca(Mock(), connect=connect)
        connect.assert_not_called()

    def test_missing_certificate_closes_connection(self):
        socket = Mock()
        socket.recv.return_value = '{"data":{"usbkey":[]}}'
        with patch.dict(os.environ, {"MMS_CA_NAME": "example", "MMS_CA_PIN": "test",
                                    "MMS_CA_POSITION": "0"}):
            with self.assertRaisesRegex(RuntimeError, "证书枚举失败"):
                login_with_ca(Mock(), connect=Mock(return_value=socket))
        socket.close.assert_called_once()

    def test_hidden_fields_support_attribute_order_quotes_and_self_closing(self):
        parser = _HiddenFields()
        parser.feed('<input value="a&amp;b" name="code"/><input name=state value=ok>')
        self.assertEqual(parser.fields, {"code": "a&b", "state": "ok"})
