# Arquitectura del proyecto

Mapa de carpetas y ficheros. Actualizar cada vez que se añada o mueva un fichero.

```
tpv-prestashop-system/
├── ARQUITECTURA.md                          # Este fichero: mapa de carpetas y ficheros
├── .gitignore                               # Excluye los Excel con datos reales (*.xlsx)
├── docs/
│   └── presupuesto/
│       └── ANALISIS_EXCEL_PRESUPUESTO.md    # Análisis de ElSuperDocumento y PRESUPUESTO_2026: entradas, proceso, fallos
└── tools/
    └── presupuesto/
        └── verificar_excels.py              # Reproduce las cifras y fallos del análisis (requiere openpyxl)
```

## Ficheros de datos (fuera del repo)

| Fichero | Descripción |
|---|---|
| `ElSuperDocumento_v19.xlsx` | Control plurianual de proyectos |
| `PRESUPUESTO_2026_v3_control_ejecucion.xlsx` | Presupuesto anual 2026 |
