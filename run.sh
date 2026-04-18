#!/bin/bash
# Script de inicialização do WFM Engine

echo ""
echo " ============================================"
echo "  WFM ENGINE - Iniciando servidor..."
echo " ============================================"
echo ""

# Instalar dependências se necessário
pip install flask numpy scipy openpyxl -q

echo ""
echo " Acesse no navegador: http://localhost:5000"
echo ""

# Iniciar o servidor
cd "$(dirname "$0")"
PYTHONPATH=./src python src/app.py "$@"
