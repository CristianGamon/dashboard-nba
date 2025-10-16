# 🏀 NBA Dashboard – Seguimiento automático de partidos y análisis de apuestas deportivas  
**(NBA Dashboard – Automated Game Tracking & Sports Betting Analysis)**

Proyecto desarrollado en **Microsoft Excel con VBA y Power Query** para automatizar el seguimiento de los partidos de la NBA y apoyar la **toma de decisiones en apuestas deportivas**, especialmente en el mercado de **total de puntos ("over/under")**.  

El dashboard permite descargar automáticamente los resultados de cada jornada, registrar las cuotas, calcular ganadores y favoritos, y analizar tendencias de anotación para identificar oportunidades de valor en las apuestas.

---

## 🚀 Funcionalidades principales  
**Main Features**
- 📅 **Descarga automática** de partidos desde [hispanosnba.com](https://www.hispanosnba.com/) mediante Power Query.  
- 🧮 **Consolidación de datos** en una base de datos interna ("BD") para análisis histórico.  
- 🏆 **Determinación automática** de ganadores, favoritos y rendimientos.  
- 🎯 **Análisis de puntos totales** por partido para orientar apuestas “Over/Under”.  
- 🧾 **Formularios interactivos** en VBA para seleccionar fechas y registrar cuotas.  
- ⚙️ Código modular y optimizado con control de errores y funciones auxiliares (`LastRow`, `SlugTeam`, etc.).  

---

## 📊 Objetivo del proyecto  
**Project Goal**

El objetivo fue transformar un archivo de Excel en una herramienta **dinámica y analítica** que automatiza la recolección de datos de la NBA, mantiene un histórico actualizado y ayuda a detectar patrones útiles para las **apuestas deportivas basadas en datos**.

---

## 🧩 Arquitectura del proyecto  
**Project Architecture**

NBA-Dashboard/
│
├── Dashboard_NBA.xlsm # Archivo principal con macros y dashboard
├── vba/
│ ├── Module1.bas # Código principal (DescargaResultados, VictoriaDerrota, etc.)
│ ├── UserForm1.frm # Formulario de cuotas (registro y validación)
│ ├── UserForm3.frm # Formulario de selección de fechas
│
├── screenshots/
│ ├── dashboard_overview.png # Vista general del panel
│ ├── form_userinput.png # Formulario interactivo
│ └── demo.gif # (opcional) animación del proceso
│
└── README.md # Este archivo
