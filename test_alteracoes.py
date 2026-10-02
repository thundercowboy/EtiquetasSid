import unittest
from types import SimpleNamespace
from unittest.mock import patch
from etiquetas_sid import App, carregar_pacientes


class AlteracoesTest(unittest.TestCase):
    def test_numero_float_sem_decimal(self):
        sheet = SimpleNamespace(iter_rows=lambda **kwargs: iter([
            ('NÚMERO VIDAS', 'NOME COMPLETO', 'DATA NASCIMENTO'),
            (673.0, 'Paciente fictício', '01/01/2000'),
        ]))
        workbook = SimpleNamespace(worksheets=[sheet], close=lambda: None)
        with patch('openpyxl.load_workbook', return_value=workbook):
            self.assertEqual(carregar_pacientes('ficticio.xlsx')['673']['numero'], '673')

    def test_aplicar_grupo_preserva_nao_selecionado(self):
        labels = [dict(numero=str(i), nome='Teste', nascimento='01/01/2000', periodo='Antigo', responsavel='Beatriz') for i in range(3)]
        tree = SimpleNamespace(selection=lambda: ('0', '2'), item=lambda *args, **kwargs: None)
        app = SimpleNamespace(labels=labels, tree=tree, period=SimpleNamespace(get=lambda: 'DIA 01 AO DIA 07'), responsible=SimpleNamespace(get=lambda: 'Amanda Altino'))
        App.save_fields(app)
        self.assertEqual(labels[0]['periodo'], 'DIA 01 AO DIA 07')
        self.assertEqual(labels[2]['responsavel'], 'Amanda Altino')
        self.assertEqual(labels[1]['periodo'], 'Antigo')
        app.period.get = lambda: ''
        app.responsible.get = lambda: 'Beatriz'
        App.save_fields(app)
        self.assertEqual(labels[0]['periodo'], 'DIA 01 AO DIA 07')
        self.assertEqual(labels[2]['responsavel'], 'Beatriz')

    def test_remover_grupo(self):
        app = SimpleNamespace(labels=['a', 'b', 'c'], tree=SimpleNamespace(selection=lambda: ('0', '2')), refresh=lambda: None)
        App.remove(app)
        self.assertEqual(app.labels, ['b'])


if __name__ == '__main__':
    unittest.main()
