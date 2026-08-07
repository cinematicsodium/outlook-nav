from types import SimpleNamespace

import pytest
from typer.testing import CliRunner

import _outlook
from _outlook import ItemType, Outlook, OutlookError
from _outlook.cli.app import app
from _outlook.services.outlook import _select_account

from .test_model_navigation import FakeAccount, FakeCollection, FakeFolder


def _accounts() -> list[_outlook.Account]:
    root = FakeFolder("Root")
    return [
        _outlook.Account(FakeAccount("First", "first@example.com", root)),
        _outlook.Account(FakeAccount("Second", "second@example.com", root)),
    ]


def test_public_package_and_cli_help_import_normally() -> None:
    assert _outlook.Outlook is Outlook
    result = CliRunner().invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "accounts" in result.stdout


def test_account_selection_handles_multiple_accounts_without_guessing() -> None:
    accounts = _accounts()
    assert _select_account(accounts) is None
    assert _select_account(accounts, "SECOND") is accounts[1]
    with pytest.raises(OutlookError):
        _select_account(accounts, "missing@example.com")


def test_close_releases_cached_com_wrappers() -> None:
    account_items = [
        FakeAccount("Only", "only@example.com", FakeFolder("Root")),
    ]
    namespace = SimpleNamespace(
        Accounts=FakeCollection(account_items),
        Class=ItemType.NAMESPACE,
    )
    client = Outlook(app=SimpleNamespace(), mapi=namespace)
    assert client.accounts

    client.close()

    assert client.account is None
    assert "accounts" not in client.__dict__
    with pytest.raises(OutlookError):
        _ = client.defaults
