from waitress import serve
from app import app

if __name__ == '__main__':
    print("Iniciando o servidor de produção local...")
    print("O sistema está disponível na rede local na porta 5000.")
    print("Acesse de outros computadores usando: http://<IP_DO_SERVIDOR>:5000")
    serve(app, host='0.0.0.0', port=5000)
