import sys
import unittest
from pathlib import Path
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from automatizar_fichadas import (
    CONFIG, detectar_duplicados_cercanos, detectar_correcciones_estado,
    generar_turnos, clasificar_turnos,
)


def df_marcas(marcas):
    return pd.DataFrame([dict(legajo='1', nombre='OPERARIO PRUEBA',timestamp=pd.Timestamp(ts),estado=est) for ts,est in marcas])

class ReglasTest(unittest.TestCase):
    def calcular(self, marcas):
        df=df_marcas(marcas)
        df, dup=detectar_duplicados_cercanos(df,CONFIG)
        df, corr=detectar_correcciones_estado(df,CONFIG)
        turnos, errores=generar_turnos(df,CONFIG)
        return clasificar_turnos(turnos,CONFIG), errores, dup, corr

    def test_limite_14(self):
        t,e,*_=self.calcular([('2026-07-02 06:00','Entrada'),('2026-07-02 20:00','Salida')])
        self.assertEqual(len(t),1)
        self.assertEqual(t.iloc[0].estado_control,'OK')

    def test_mas_de_14_hasta_17(self):
        t,e,*_=self.calcular([('2026-07-02 06:00','Entrada'),('2026-07-02 22:00','Salida')])
        self.assertEqual(len(t),1)
        self.assertIn('REVISAR',t.iloc[0].estado_control)

    def test_mas_de_17_excluido(self):
        t,e,*_=self.calcular([('2026-07-02 06:00','Entrada'),('2026-07-03 00:00','Salida')])
        self.assertTrue(t.empty)
        self.assertIn('Turno crítico no computado',list(e.tipo_error))

    def test_entrada_muy_antigua_reemplazada(self):
        t,e,*_=self.calcular([('2026-07-07 16:55','Entrada'),('2026-07-13 05:56','Entrada'),('2026-07-13 17:03','Salida')])
        self.assertEqual(len(t),1)
        self.assertAlmostEqual(float(t.iloc[0].horas),11+7/60,places=2)
        self.assertIn('Entrada abierta demasiado antigua',list(e.tipo_error))

    def test_duplicado_cercano(self):
        t,e,dup,_=self.calcular([('2026-07-02 06:00','Entrada'),('2026-07-02 06:01','Entrada'),('2026-07-02 17:00','Salida')])
        self.assertEqual(len(dup),1)
        self.assertEqual(len(t),1)

    def test_salida_inicial(self):
        t,e,*_=self.calcular([('2026-07-01 05:52','Salida')])
        self.assertEqual(len(t),1)
        self.assertAlmostEqual(float(t.iloc[0].horas),5+52/60,places=2)

    def test_corte_quincena_abierto(self):
        t,e,*_=self.calcular([('2026-07-15 19:00','Entrada')])
        self.assertEqual(len(t),1)
        self.assertTrue(t.iloc[0].turno_recortado_por_cierre_quincena=='Sí')

    def test_correccion_inmediata(self):
        t,e,dup,corr=self.calcular([('2026-07-02 06:00','Salida'),('2026-07-02 06:01','Entrada'),('2026-07-02 17:00','Salida')])
        self.assertEqual(len(corr),1)
        self.assertEqual(len(t),1)

if __name__=='__main__':
    unittest.main()
