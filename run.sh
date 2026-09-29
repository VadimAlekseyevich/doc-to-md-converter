#!/usr/bin/env bash
# Bootstrap, install only missing requirements, build editable package and launch.
set -Eeuo pipefail
CALLER_CWD="$PWD"
cd -- "$(dirname -- "${BASH_SOURCE[0]}")"

info() { printf '[setup] %s\n' "$*"; }
fail() { printf '[setup] Ошибка: %s\n' "$*" >&2; exit 1; }

as_admin() {
    if (( EUID == 0 )); then "$@"
    elif command -v sudo >/dev/null 2>&1; then sudo "$@"
    else fail "Требуются права администратора для установки: $*. Установите пакет вручную и повторите запуск."
    fi
}

install_os_packages() {
    if command -v apt-get >/dev/null 2>&1; then
        as_admin apt-get update
        as_admin apt-get install -y "$@"
    elif command -v dnf >/dev/null 2>&1; then
        as_admin dnf install -y "$@"
    elif command -v brew >/dev/null 2>&1; then
        brew install "$@"
    else
        fail "Автоустановка не поддерживается этой ОС. Установите Python 3.11+ и Tkinter вручную."
    fi
}

python_works() {
    command -v "$1" >/dev/null 2>&1 &&
        "$1" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)' >/dev/null 2>&1
}

find_python() {
    local candidate
    for candidate in python3 python3.13 python3.12 python3.11 python; do
        if python_works "$candidate"; then command -v "$candidate"; return 0; fi
    done
    return 1
}

install_python() {
    info 'Python 3.11+ не обнаружен, пытаюсь установить.'
    if command -v apt-get >/dev/null 2>&1; then
        as_admin apt-get update
        local minor
        for minor in 13 12 11; do
            if apt-cache show "python3.$minor" >/dev/null 2>&1 && apt-cache show "python3.$minor-venv" >/dev/null 2>&1; then
                as_admin apt-get install -y "python3.$minor" "python3.$minor-venv"
                return
            fi
        done
        as_admin apt-get install -y python3 python3-venv
    elif command -v dnf >/dev/null 2>&1; then
        install_os_packages python3
    elif command -v brew >/dev/null 2>&1; then
        install_os_packages python
    else
        fail 'Установите Python 3.11+ вручную и повторите запуск.'
    fi
}

PYTHON="$(find_python || true)"
if [[ -z "$PYTHON" ]]; then
    install_python
    PYTHON="$(find_python || true)"
    [[ -n "$PYTHON" ]] || fail 'Пакетный менеджер не предоставил Python 3.11+. Установите совместимую версию вручную.'
fi
info "Используется $("$PYTHON" --version) ($PYTHON)"

if [[ ! -x .venv/bin/python ]] || ! .venv/bin/python -c 'import sys; sys.exit(sys.version_info < (3, 11))' >/dev/null 2>&1; then
    info 'Создаю локальное окружение .venv (с доступом к уже установленным пакетам).'
    if ! "$PYTHON" -m venv --system-site-packages .venv; then
        if command -v apt-get >/dev/null 2>&1; then
            venv_pkg="$($PYTHON -c 'import sys; print(f"python{sys.version_info.major}.{sys.version_info.minor}-venv")')"
            if apt-cache show "$venv_pkg" >/dev/null 2>&1; then
                install_os_packages "$venv_pkg"
            else
                install_os_packages python3-venv
            fi
            "$PYTHON" -m venv --system-site-packages .venv || fail 'Не удалось создать .venv (для отдельных версий Python нужен пакет python3.X-venv).'
        else
            fail 'Не удалось создать .venv; установите модуль venv для используемого Python.'
        fi
    fi
fi
VENV_PYTHON="$PWD/.venv/bin/python"
# A nested venv does not inherit its parent's site-packages, even with
# --system-site-packages. Reuse packages from an already active Python venv
# instead of downloading the same wheels again.
if "$PYTHON" -c 'import sys; sys.exit(0 if sys.prefix != sys.base_prefix else 1)' >/dev/null 2>&1 &&
   [[ "$($PYTHON -c 'import sys; print(sys.version_info[:2])')" == "$($VENV_PYTHON -c 'import sys; print(sys.version_info[:2])')" ]]; then
    local_site="$($VENV_PYTHON -c 'import sysconfig; print(sysconfig.get_path("purelib"))')"
    "$PYTHON" -c 'import site, sys, os; print("\n".join(p for p in site.getsitepackages() if os.path.isdir(p) and os.path.commonpath([os.path.realpath(p), os.path.realpath(sys.prefix)]) == os.path.realpath(sys.prefix)))' > "$local_site/_preexisting_venv.pth"
    info 'Использую также уже установленные библиотеки исходного Python-окружения.'
fi
"$VENV_PYTHON" -m pip --version >/dev/null 2>&1 || "$VENV_PYTHON" -m ensurepip --upgrade

# Check version constraints, not just import presence. Each installation is conditional.
if ! "$VENV_PYTHON" -c 'from importlib.metadata import version; import re; m=re.match(r"^(\d+)\.(\d+)", version("python-docx")); assert m and (1, 1) <= tuple(map(int, m.groups())) < (2, 0)' 2>/dev/null; then
    info 'Устанавливаю/обновляю python-docx (требуется >=1.1,<2).'
    "$VENV_PYTHON" -m pip install 'python-docx>=1.1,<2'
else
    info 'python-docx уже подходит.'
fi
if ! "$VENV_PYTHON" -c 'from importlib.metadata import version; import re; m=re.match(r"^(\d+)\.(\d+)", version("lxml")); assert m and (4, 9) <= tuple(map(int, m.groups())) < (7, 0)' 2>/dev/null; then
    info 'Устанавливаю/обновляю lxml (требуется >=4.9,<7).'
    "$VENV_PYTHON" -m pip install 'lxml>=4.9,<7'
else
    info 'lxml уже подходит.'
fi
if ! "$VENV_PYTHON" -c 'from importlib.metadata import version; import re; m=re.match(r"^(\d+)", version("setuptools")); assert m and int(m.group(1)) >= 68' 2>/dev/null; then
    info 'Устанавливаю setuptools для сборки.'
    "$VENV_PYTHON" -m pip install 'setuptools>=68'
fi

# Editable install uses local source; --no-deps and --no-build-isolation forbid
# re-downloading dependencies on subsequent launches.
info 'Собираю и регистрирую приложение в .venv (без скачивания зависимостей).'
"$VENV_PYTHON" -m pip install --quiet --no-deps --no-build-isolation -e .

if (( $# == 0 )); then
    if ! "$VENV_PYTHON" -c 'import tkinter' >/dev/null 2>&1; then
        info 'Tkinter не обнаружен, пытаюсь установить системный пакет.'
        if command -v apt-get >/dev/null 2>&1; then
            install_os_packages python3-tk
        elif command -v dnf >/dev/null 2>&1; then
            install_os_packages python3-tkinter
        elif command -v brew >/dev/null 2>&1; then
            install_os_packages python-tk
        else
            fail 'Установите Tkinter для GUI вручную или запустите ./run.sh path/to/file.docx в CLI-режиме.'
        fi
        "$VENV_PYTHON" -c 'import tkinter' >/dev/null 2>&1 || fail 'Tkinter всё ещё недоступен для выбранной версии Python. Используйте CLI или установите соответствующий tkinter.'
    fi
fi
info 'Запускаю конвертер.'
cd -- "$CALLER_CWD"  # CLI path arguments remain relative to the launching shell.
exec "$VENV_PYTHON" -m doc_to_md_converter "$@"
