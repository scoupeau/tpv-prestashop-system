# Análisis de los Excel de presupuesto (fase previa a programar)

Fecha: 2026-10-01 · Estado: análisis, sin código de aplicación todavía.

Ficheros analizados (no se versionan en el repo, ver `.gitignore`):

| Fichero | Rol |
|---|---|
| `ElSuperDocumento_v19.xlsx` | Control **plurianual** de proyectos (2022-2033): catálogo, cobros, ejecución, imputación de ayudas |
| `PRESUPUESTO_2026_v3_control_ejecucion.xlsx` | Presupuesto **anual** 2026 + cash flow + plan de actuación del Protectorado |

Las cifras de este documento se pueden reproducir con:

```bash
pip install openpyxl
python3 tools/presupuesto/verificar_excels.py <ElSuperDocumento.xlsx> <PRESUPUESTO_2026.xlsx>
```

---

## 1. Estructura de cada libro

### ElSuperDocumento_v19.xlsx
La **fila** es la clave: el proyecto de la fila *n* de `PROYECTOS` está en la fila *n* de las demás hojas.

| Hoja | Contenido |
|---|---|
| `PROYECTOS` | Maestro: código, responsable, fechas, observaciones, estado ("SIGUE EN 2026", "FINALIZADO Y COBRADO"), fecha de traspaso de ficha |
| `plurianuales_ingresos` | Calendario **mensual** de cobros de subvenciones/convenios plurianuales (2022-2033) |
| `anuales_ingresos` | Anuales y donaciones: cantidad a ejecutar, concedida, fecha justificación + calendario (desplazado 3 columnas) |
| `facturables` | Calendario de facturación |
| `INGRESOS TOTALES` | Suma de las tres anteriores por mes |
| `control_plurianuales` | Gasto ejecutado por trimestre → % ejecución → ayuda imputable a resultados (ctas. 740/723) |
| `PREV_PERSONAL` / `PREV_MATERIAL` | Gasto previsto por trimestre (2022-2025) y anual (2026-2029) |
| `EJE_PERSONAL` / `EJE_MATERIAL` | Gasto ejecutado por trimestre (hasta 4T 2025) |
| `GRAFICOS` | Liquidez trimestral (solo 2023-2024) |
| Ocultas: `NOMINAL`, `PREUBA`, `Hoja3`, `Hoja6`, `sub-conv plurianuales` | Versiones antiguas / tablas dinámicas |

### PRESUPUESTO_2026_v3_control_ejecucion.xlsx

| Hoja | Contenido |
|---|---|
| `PPTO CON PROP` | Presupuesto de ingresos y gastos (devengo) por proyecto |
| `CASH FLOW` | Misma estructura, pero ingresos = cobros esperados en el año; saldo inicial → saldo final |
| `PROTEC CON PROP` | Plan de actuación en formato Protectorado (recursos a emplear / obtener por actividad). **Es el entregable** |
| `GASTOS GENERALES` | Gastos de estructura por cuenta contable con histórico 2021-2026 y comentarios |
| `GASTOS No Imputables` | Auditoría, bancos, publicidad, gastos financieros, asociaciones |
| `FACT EXTRA`, `DONACION`, `CONVENIO`, `FONDO NAT` | Ingresos no previstos registrados durante el año (control de ejecución) |

## 2. Entradas de datos y qué se hace con cada una

| Entrada | Dónde | Uso |
|---|---|---|
| Alta de proyecto (código, nombre, responsable, fechas) | `PROYECTOS` B-G | Todas las hojas lo enlazan con `=PROYECTOS!Bn` |
| Estado del proyecto | `PROYECTOS` H, L | Decide (a mano) si entra en el presupuesto del año siguiente |
| Gasto previsto por proyecto (de la ficha económica) | `PREV_*` → `PPTO CON PROP` G-AA | Desglose en 21 conceptos: personal, viajes, asistencias externas, inventariable, tierras… |
| Agrupación por cuenta contable | `PPTO` AB-AN | **Fórmula distinta en cada fila**, escrita a mano (`AC4=+I4+L4+O4+P4+Q4+R4+Z4`, `AK7=+X7`…) |
| Overheads | `PPTO` N | No es gasto; solo entra en el cálculo del ingreso `(gastos+N) × % cofinanciación` |
| Ingreso por tipo | `PPTO` AO subvención · AP convenio · AQ donación · AR facturación | Subvenciones: % sobre gasto. Varios financiadores → sub-filas (LIFE AWOM 7, LIFE Ballesteros 3) |
| Gastos generales por cuenta | `GASTOS GENERALES` col. O | Alimentan la fila 2101 "Coordinación general" |
| Gastos no imputables | `GASTOS No Imputables` col. O | Alimentan la fila 3101 |
| **Masa salarial total 2.636.195,80 €** (escrita dentro de la fórmula) | `PPTO!G103` | `G103 = 2.636.195,80 + otros gastos sociales − personal Act.1 − personal Act.2` → personal no asignado a proyectos = estructura (749.359,53 €). **Palanca principal del cuadre** |
| Reparto de comunes | `PPTO` filas 107/109 | 94 % Actividad 1 / 6 % Actividad 2 |
| Ingresos no previstos | `FACT EXTRA`, `DONACION`, `CONVENIO`, `FONDO NAT` | Suben a las filas 84/99 del presupuesto |
| Cobros del año | `CASH FLOW` AO-AR | 42 de 76 códigos coinciden exactamente con `INGRESOS TOTALES` 2026; el resto, retocado a mano |
| Saldo inicial de tesorería 3.627.020,93 € | `CASH FLOW!AO116` | Saldo inicial + cobros − pagos = saldo final 4.486.052,14 € |
| Gasto ejecutado por trimestre (contabilidad) | `control_plurianuales` AG-AJ, `EJE_*` | % ejecución e ingreso imputable. **2026 casi vacío (1 dato)** |

## 3. Proceso que sigue el usuario (inferido)

1. Revisar `PROYECTOS`: qué sigue, qué termina, qué es nuevo, fechas.
2. Para cada proyecto vivo, sacar el gasto anual de su ficha económica, pasarlo a `PREV_*` (solo proyectos nuevos) y teclearlo por concepto en `PPTO CON PROP`.
3. Calcular el ingreso según el tipo de financiación (% sobre gasto o importe fijo); sub-filas si hay varios financiadores.
4. Actualizar `GASTOS GENERALES` y `GASTOS No Imputables` (columna nueva por año, a partir del histórico y con comentarios).
5. Cuadrar personal: poner la masa salarial total (de la hoja de salarios externa); la diferencia con lo asignado va a coordinación.
6. Comprobar `G124 = 0` (total por naturaleza = total por cuentas). Hoy da 0.
7. Mirar el resultado: ingresos 6.361.556,54 € frente a gastos 5.774.962,60 € + inversiones 1.444.361,50 € (total recursos 7.219.324,10 €). No hay una celda de resultado explícita: el ajuste se hace moviendo renovaciones supuestas y la masa salarial.
8. Replicar en `CASH FLOW` con los cobros del SuperDocumento para comprobar la liquidez.
9. `PROTEC CON PROP` se genera solo y se entrega al Protectorado.
10. Durante el año: registrar ingresos extra y la ejecución por trimestre.

## 4. Fallos y riesgos verificados

| # | Hallazgo | Ubicación |
|---|---|---|
| 1 | 4104 HORIZON PES REWARD tiene 25.892,08 €, la misma cifra que 4101; el SuperDocumento dice 161.365,31 € para 2026 → posible copia errónea | `CASH FLOW!AO90` |
| 2 | La fila de total suma hasta la fila 82 en G, pero hasta la 81 de H a AO → los gastos no personales de CON DATA (1702) quedan fuera | `CASH FLOW` fila 85 |
| 3 | Reparto de comunes inconsistente: 94/6 en PPTO y 91/9 en PROTEC | `PPTO` 107/109 vs `PROTEC` col. D/G |
| 4 | Masa salarial y saldo inicial escritos dentro de fórmulas / sin celda de parámetro | `PPTO!G103`, `CASH FLOW!AO116` |
| 5 | Fechas guardadas como texto: `01/10/202`, `…....`, `xx/xx/2026`, `xx/xx/xxxx`, `sin fin`, `031/12/2025` | `PPTO` E8, F24, E34, F34, F57, F67 |
| 6 | Código 1202 en 3 filas; Villacañas 2025-26 y 2026-27 llevan cada una un año completo de coste (¿duplica 2026?) | `PPTO` filas 36-38 |
| 7 | Proyectos con fecha fin pasada e importes en 2026 (p. ej. Nestlé Cereales, fin 31/08/2024, 170.000 €) | `PPTO` fila 92 |
| 8 | La cabecera dice "2025" pero alimenta 2026 | `GASTOS No Imputables!O1` |
| 9 | 266 filas desalineadas una posición respecto a `PROYECTOS` | `EJE_PERSONAL`, `EJE_MATERIAL` |
| 10 | 21 vínculos externos a fichas en `\\192.168.10.251\...`; no funcionan fuera de esa red | SuperDocumento |
| 11 | Hojas sin actualizar: `INGRESOS TOTALES!E` suma solo 2022-2023; `GRAFICOS` llega a 2024 | SuperDocumento |
| 12 | Añadir un proyecto obliga a insertar la misma fila en ~8 hojas | SuperDocumento |
| 13 | Diferencias de cobro destacadas: 1117 PRODIEL (SuperDoc 645.933 € / CASH FLOW 0), 1221 LIFE Ballesteros y 1222 LIFE HumedalES (CASH FLOW 786.935 € y 620.250 € / SuperDoc 0) | `CASH FLOW` vs `INGRESOS TOTALES` |

## 5. Preguntas abiertas antes de diseñar

1. ¿Qué significa "cuadrar": resultado = 0, ingresos ≥ gastos, saldo de tesorería mínimo, o personal asignado + estructura = masa salarial?
2. ¿De dónde sale la masa salarial total? ¿Existe una hoja de salarios por persona que deba ser entrada?
3. ¿El reparto de comunes correcto es 94/6 o 91/9?
4. ¿El presupuesto del año siguiente parte de copiar el anterior o de los proyectos marcados "SIGUE" en el SuperDocumento?
5. ¿Las fichas económicas externas siguen siendo la fuente del gasto por proyecto?

## 6. Propuesta de flujo (borrador)

- **Maestro único de proyectos** (código como clave, no la fila).
- **Una entrada por proyecto y año**: gasto por concepto + lista de financiadores con su % o importe.
- **Tabla de mapeo concepto → cuenta contable** común (sustituye las fórmulas fila a fila).
- **Parámetros** del ejercicio: masa salarial, saldo inicial, reparto entre actividades.
- **Salidas generadas**: PPTO, CASH FLOW y PROTEC, con comprobaciones automáticas de cuadre.
