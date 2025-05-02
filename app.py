import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO

# Conexión a SQL Server
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=192.168.1.101;'
    'DATABASE=Sli;'
    'UID=reportes;'
    'PWD=reportes01;'
)

# Título de la aplicación
st.title("Reporte TECNOGLOBAL")

# Entrada de rango de fechas
fecha_inicio = st.date_input("Fecha de inicio")
fecha_fin = st.date_input("Fecha de fin")

# Botón para consultar
if st.button("Consultar"):
    with st.spinner("Consultando datos..."):
        query = f"""
        SELECT 
            DinCab.Despacho, DinCab.NumIdentif, CONVERT(VARCHAR(10), Dincab.FecAcep, 103) as Fecha, Ape.Ref_Cliente,
            DinCab.Consignatario, DinItems.Item, DinItems.CIP, DinItems.Arancel, DinItems.Mercancia, DinItems.Marca, DinItems.Variedad,
            DinItems.CtdaMercancias, 
            UniMed.Descripcion + ' (' + Convert(varchar(5), DinItems.Id_UnidMedida) + ')' As UniMed,
            DinCab.Cliente
        FROM Sli..Din_Cab DinCab
        INNER JOIN Sli..Apertura Ape ON DinCab.Despacho = Ape.Despacho
        INNER JOIN Sli..Din_Items DinItems ON DinCab.Despacho = DinItems.Despacho
        LEFT JOIN Sli..UnidMedida UniMed ON DinItems.Id_UnidMedida = UniMed.Id_Adu_UnidMed
        WHERE DinCab.Id_Cliente IN('968230204', '836281004', '867312005')
            AND Dincab.FecAcep BETWEEN ? AND ?
            AND Ape.Nulo = 0
            AND DinItems.Arancel IN('84433214', '84433110', '84714900', '84713040', '84713030', '84715000', '84433217')
        ORDER BY DinCab.Cliente, DinCab.FecAcep
        """
        df = pd.read_sql(query, conn, params=[fecha_inicio, fecha_fin])
        st.success(f"{len(df)} filas encontradas.")
        st.dataframe(df)

        # Exportar a Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Reporte')

        st.download_button(
            label="Descargar en Excel",
            data=output.getvalue(),
            file_name="reporte_despachos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
