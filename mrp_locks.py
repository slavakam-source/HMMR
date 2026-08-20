"""Заглушка для отсутствующего в этом окружении модуля mrp_locks.

Оригинальный модуль реализует файловую блокировку "один расчёт MRP
на папку одновременно" (используется GUI-обвязкой). В этом
изолированном контейнере параллельных запусков не бывает, поэтому
блокировка — no-op. Диагностический прогон движка, сам движок не
менялся.
"""
import os


def run_lock_path(root):
    return os.path.join(root, '.mrp_run.lock')


def active_lock_message(root, out_xlsx):
    return None


def try_acquire(lock_path, holder_name):
    return True, {'holder': holder_name, 'pid': os.getpid()}


def format_lock_holder(info):
    return str(info)


def release(lock_path, pid):
    return None
