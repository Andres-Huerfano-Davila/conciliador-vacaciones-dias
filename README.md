# Conciliación de vacaciones

Aplicación Streamlit independiente para conciliar exclusivamente días. No calcula costos ni importes monetarios. No modifica los libros ni las automatizaciones existentes.

## Ejecutar

Instale las dependencias con `python -m pip install -r requirements.txt` y abra `run.bat`, o ejecute `python -m streamlit run app.py` desde esta carpeta.

## Entradas

SAP: masterdata, ausentismos, cuotas IT2006 y provisiones PVVA, TXT no convertido (pipe o tabulado; también CSV y XLSX). SF: CSV de movimientos de saldo y solicitudes pendientes. Detecta UTF-8, UTF-16 y Windows-1252, encabezados SAP y metadatos SF. Un XLS que realmente contiene texto se lee por contenido. XLS binario antiguo requiere exportarlo como no convertido.

Columnas mínimas: identificador SAP en cada fuente; masterdata `Área de nómina` y `Fecha`/`Fecha de ingreso`/`Alta`; ausencias `Cl.`/`ClAbPre`, `Inicio`, `Final`, `Días`; IT2006 `Cantidad`, `LiquidCont`; provisión `TipoCalc`, `Dias Sd A`; SF saldo `User ID`, `Amount posted` y, para movimientos, `Posting Date`; pendientes `ID personal`, `approvalStatus`, `quantityInDays`. Se conservan las primeras columnas ante encabezados duplicados de SAP.

## Reglas

Saldo histórico, contrastado con el motor local `nomina_app/consolidated_engine.py`: días efectivos desde ingreso/reingreso × multiplicador anual / divisor, menos vacaciones disfrutadas, con mínimo cero. ZM, Administrativos y SENA Productivo: meses comerciales / 360. ZL/Mensual Admon 365, ZH/Parcial Hora y ZP/Parcial Día: calendario / 365. Multiplicador 15 o 22 para áreas de vacaciones 22. La regla mensual de reemplazo de proyecciones negativas (divisor 360 para todas las áreas) no se aplica al saldo histórico.

SENA Productivo se identifica específicamente por `Área de personal`: anula la regla 365 del área de nómina y aplica calendario comercial / 360. Para ausencias no remuneradas cuya duración original no supera 30 días, agrega el día 31 comprendido en el tramo; las ausencias largas no reciben ese agregado. Mantiene una columna explícita de identificación SENA.

No remuneradas: códigos 0189,0191,0195,0196,0197,0380,0381,0397,0399, editables. Se fusionan intervalos para no descontar doble. Part time hora descuenta ausencias en lunes a viernes; part time día, fines de semana y lunes festivos. Vacaciones 0100: respeta días SAP cuando la ausencia está enteramente dentro del período; reconstruye tramos parciales por calendario del área. Calendario festivo replica el motor existente, incluyendo su regla desde 2026.

SAP 2006 = Cantidad − LiquidCont. Provisión = suma PVVA / Dias Sd A (primera columna de ese nombre). SF = suma de Amount posted para CO_vacatio hasta el corte, con opción explícita para ajustes manuales posteriores. Pendientes PENDING se informan sin descontarlos nuevamente. Una ausencia de saldo permanece sin información. Se incluyen empleados presentes únicamente en otras fuentes.

Umbrales estrictos: revisar si diferencia absoluta > 0,25 en cálculo/provisión o > 1,25 en los otros tres cruces. Igual al límite es OK. Se compara sin redondear; se muestran dos decimales.

## Condiciones y límites

Confirme que el histórico cubre todo el vínculo laboral; sin confirmación no se emite cálculo independiente. La fecha de alta técnica del maestro debe representar el inicio del vínculo vigente. Si hubo reingreso, suministre `Fecha reingreso` o una fecha de ingreso actualizada. Un maestro con datos contradictorios por empleado queda en revisión. Un maestro actual no permite reconstruir cambios históricos de área. Esta primera versión no admite saldo inicial para historias recortadas ni historia salarial para reconstruir ausencias part time con cambios de jornada. Use los días SAP para ausencias completas. Los saldos SAP exportados deben corresponder al corte: la aplicación no reconstruye el consumo histórico de cuotas de una descarga actual.

Salida: un reporte en pantalla con filtros y una sola hoja Excel. La evidencia JSON opcional contiene nombres, conteos y hashes de los archivos, sin adjuntar las bases.

## GitHub y Streamlit

Suba únicamente código, README y requirements; nunca bases de empleados. En Streamlit configure `app.py` como archivo principal. Para datos de nómina utilice un despliegue con acceso corporativo autorizado. El repositorio contiene código; las bases se cargan en cada ejecución.

## Verificar

`python -m unittest discover -s tests -v`
