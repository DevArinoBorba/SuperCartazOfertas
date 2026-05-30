import os
import signal
import subprocess
import sys
import time


SERVER_SCRIPT = os.path.join(os.path.dirname(__file__), "run_server.py")


def start_server():
    creationflags = 0
    if os.name == "nt":
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP

    return subprocess.Popen(
        [sys.executable, SERVER_SCRIPT],
        cwd=os.path.dirname(__file__),
        creationflags=creationflags,
    )


def stop_server(process):
    if process.poll() is not None:
        return

    try:
        if os.name == "nt":
            os.kill(process.pid, signal.CTRL_BREAK_EVENT)
        else:
            process.terminate()
        process.wait(timeout=8)
    except Exception:
        process.kill()
        process.wait(timeout=5)


def print_commands():
    print()
    print("Comandos disponiveis nesta janela:")
    print("  R + Enter = reiniciar servidor")
    print("  S + Enter = sair e encerrar servidor")
    print()


def main():
    print("=" * 46)
    print("Gerenciador do Servidor de Cartazes")
    print("=" * 46)
    print()

    process = start_server()
    print_commands()

    try:
        while True:
            if process.poll() is not None:
                print()
                print("O servidor foi encerrado.")
                choice = input("Digite R para iniciar novamente ou S para sair: ").strip().lower()
            else:
                choice = input("Comando: ").strip().lower()

            if choice == "r":
                print()
                print("Reiniciando servidor...")
                stop_server(process)
                time.sleep(1)
                process = start_server()
                print_commands()
            elif choice == "s":
                print()
                print("Encerrando servidor...")
                stop_server(process)
                break
            elif not choice:
                continue
            else:
                print("Comando invalido. Use R para reiniciar ou S para sair.")
    except KeyboardInterrupt:
        print()
        print("Encerrando servidor...")
        stop_server(process)


if __name__ == "__main__":
    main()
