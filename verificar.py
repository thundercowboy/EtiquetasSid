"""Verificação com dados fictícios, sem expor o banco real."""
import tempfile
from pathlib import Path
from openpyxl import Workbook
from etiquetas_sid import normalizar_numero, carregar_pacientes, gerar_pdf, DEFAULT_CONFIG

assert normalizar_numero('000123') == normalizar_numero(123)
assert normalizar_numero('000') == '0'
with tempfile.TemporaryDirectory() as folder:
    path = Path(folder) / 'BD.xlsx'
    wb = Workbook()
    wb.active.append(['NÚMERO VIDAS', 'NOME COMPLETO', 'DATA NASCIMENTO'])
    wb.active.append([123, 'Paciente Fictício', '14/03/1958'])
    wb.save(path)
    patients = carregar_pacientes(path)
    assert patients['123']['nascimento'] == '14/03/1958'
    wb.active.append(['000123', 'Outro Paciente', '01/01/2000'])
    wb.save(path)
    try:
        carregar_pacientes(path)
    except ValueError as exc:
        assert 'duplicado' in str(exc)
    else:
        raise AssertionError('Duplicidade não detectada')
output = Path('tmp/pdfs')
output.mkdir(parents=True, exist_ok=True)
label = dict(numero='123', nome='PACIENTE FICTÍCIO PARA TESTE', nascimento='14/03/1958', periodo='DIA 01 AO DIA 07', responsavel='Amanda Altino')
gerar_pdf(output / 'teste.pdf', [label] * 9, DEFAULT_CONFIG)
print('Busca, datas, duplicidade e geração de 9 etiquetas: OK')
