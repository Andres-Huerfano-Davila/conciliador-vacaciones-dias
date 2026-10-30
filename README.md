# Conciliación de vacaciones

Aplicación Streamlit independiente para conciliar exclusivamente días. No calcula costos ni importes monetarios. No modifica los libros ni las automatizaciones existentes.

## Ejecutar

Instale las dependencias con `python -m pip install -r requirements.txt` y abra `run.bat`, o ejecute `python -m streamlit run app.py` desde esta carpeta.

## Entradas

Siete cargas obligatorias y una opcional: masterdata, AUSENTISMOS_FULLT.txt, AUSNOM_PART_TIME.txt, AUS_PT_2026.txt (opcional; si se carga, se une al histórico part time), cuotas IT2006 y provisiones PVVA, TXT no convertido (pipe o tabulado; también CSV y XLSX). SF: CSV de movimientos de saldo y solicitudes pendientes. Detecta UTF-8, UTF-16 y Windows-1252, encabezados SAP y metadatos SF. Un XLS que realmente contiene texto se lee por contenido. XLS binario antiguo requiere exportarlo como no convertido.

Columnas mínimas: identificador SAP en cada fuente; masterdata `Área de nómina`, `Área de personal` y `Fecha`/`Fecha de ingreso`/`Alta`; ausencias `Cl.`/`ClAbPre`, `Inicio`, `Final`, `Días`; IT2006 `Cantidad`, `LiquidCont`; provisión `TipoCalc`, `Dias N Sld`; SF saldo `User ID`, `Amount posted` y, para movimientos, `Posting Date`; pendientes `ID personal`, `approvalStatus`, `quantityInDays`. Se retienen únicamente columnas necesarias para conciliar y se permiten cargas de hasta 1 GB por archivo. Se toleran columnas finales vacías omitidas solo cuando ninguna columna usada en el cruce está ausente.

Las ausencias full time se aplican a ZM/ZL y SENA Productivo. Para ZH/ZP se unen los dos archivos part time usando `Tot.días nóm.`. Se eliminan duplicados por empleado, clase, fechas y cantidad; cantidades contradictorias para el mismo período dejan el cálculo independiente sin información. La evidencia conserva los conteos, nombres y hashes de las tres fuentes. La cobertura histórica debe abarcar ambas poblaciones desde ingreso/reingreso.

## Reglas

Saldo histórico, contrastado con el motor local `nomina_app/consolidated_engine.py`: días efectivos desde ingreso/reingreso × multiplicador anual / divisor, menos vacaciones disfrutadas, con mínimo cero. ZM, Administrativos y SENA Productivo: meses comerciales / 360. ZL/Mensual Admon 365, ZH/Parcial Hora y ZP/Parcial Día: calendario / 365. Multiplicador 15 o 22 para áreas de vacaciones 22. La regla mensual de reemplazo de proyecciones negativas (divisor 360 para todas las áreas) no se aplica al saldo histórico.

SENA Productivo se identifica específicamente por `Área de personal`: anula la regla 365 del área de nómina y aplica calendario comercial / 360. Para ausencias no remuneradas cuya duración original no supera 30 días, agrega el día 31 comprendido en el tramo; las ausencias largas no reciben ese agregado. Mantiene una columna explícita de identificación SENA.

No remuneradas: códigos 0189,0191,0195,0196,0197,0380,0381,0397,0399, editables. Se fusionan intervalos para no descontar doble. Part time hora descuenta ausencias en lunes a viernes; part time día, fines de semana y lunes festivos. Vacaciones 0100: respeta días SAP cuando la ausencia está enteramente dentro del período; reconstruye tramos parciales por calendario del área. Calendario festivo replica el motor existente, incluyendo su regla desde 2026.

SAP 2006 = Cantidad − LiquidCont. Provisión = suma de `Dias N Sld` de PVVA (saldo nuevo). `Dias Sd A` es el saldo anterior y nunca se usa como sustituto del saldo nuevo. SF = suma de Amount posted para CO_vacatio hasta el corte, con opción explícita para ajustes manuales posteriores. Pendientes PENDING y vacaciones futuras APPROVED se informan por separado, sin descontarlos nuevamente ni normalizar automáticamente saldos de sistemas.

El reporte principal usa empleados del maestro con ingreso no posterior al corte. Identificadores presentes solo en fuentes históricas se cuentan en trazabilidad; se incluyen únicamente al activar la opción correspondiente. Cada comparación presenta sus propios conteos OK/Revisar/Sin información. El estado global es Revisar si alguna diferencia numérica supera el umbral; Sin información si no hay diferencias fuera del umbral pero falta algún cruce; OK únicamente si los cuatro cruces están completos y dentro del umbral.

Tolerancias definidas por el usuario en Reglas y opciones, una por cruce (por defecto 0,25 en cálculo/provisión y 1,25 en los otros tres). Revisar si la diferencia absoluta supera la tolerancia. Igual al límite es OK. Se compara sin redondear; se muestran dos decimales.

## Condiciones y límites

La app toma como inicio del histórico la primera ausencia de los archivos cargados; en Opciones avanzadas se puede indicar la fecha "desde" del Query SAP. Quien ingresó antes de esa fecha queda sin cálculo independiente. Ingresos anteriores al inicio del histórico requieren historia adicional o una base inicial: no se calcula un saldo inflado suponiendo que no hubo disfrutes anteriores. Los cruces entre sistemas se conservan. La fecha de alta técnica del maestro debe representar el inicio del vínculo vigente. Si hubo reingreso, suministre `Fecha reingreso` o una fecha de ingreso actualizada. Un maestro actual no permite reconstruir cambios históricos de área. Esta versión no admite saldo inicial para historias recortadas ni historia salarial para reconstruir ausencias part time con cambios de jornada. Los saldos SAP exportados deben corresponder al corte: la aplicación no reconstruye consumo histórico de cuotas de una descarga actual.

Salida: un reporte en pantalla con filtros y una sola hoja Excel. La evidencia JSON opcional contiene nombres, conteos y hashes de los archivos, sin adjuntar las bases.

## GitHub y Streamlit

Suba únicamente código, README y requirements; nunca bases de empleados. En Streamlit configure `app.py` como archivo principal. Para datos de nómina utilice un despliegue con acceso corporativo autorizado. El repositorio contiene código; las bases se cargan en cada ejecución.

## Verificar

`python -m unittest discover -s tests -v`
