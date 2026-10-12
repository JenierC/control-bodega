if st.button("🚀 Enviar Pedimento Completo a Bodega", type="primary"):
                conn = sqlite3.connect("database.db")
                cursor = conn.cursor()
                
                ids_guardados = []
                for item in st.session_state['carrito']:
                    cursor.execute('''
                        INSERT INTO solicitudes (funcionario_id, articulo_id, cantidad, es_valido, justificacion, estado)
                        VALUES (?, ?, ?, ?, ?, 'PENDIENTE')
                    ''', (st.session_state['func_id'], item['codigo'], item['cantidad'], 1, justificacion_global))
                    ids_guardados.append(cursor.lastrowid)
                conn.commit()
                
                hash_temp = hashlib.sha256(str(datetime.now()).encode()).hexdigest()
                consecutivo, ruta_excel = generar_formulario_oficial_bys(ids_guardados, "Departamento de Servicios Generales", "Pendiente de Firma", hash_temp)
                ruta_pdf = ruta_excel.replace(".xlsx", ".pdf")
                exportar_excel_a_pdf(ruta_excel, ruta_pdf)
                conn.close()

                # ENVÍO AUTOMÁTICO DE CORREO A AGONZALEZ CON COPIA A GGUTIERREZ Y JCAMACHO
                enviar_correo_con_pdf_adjunto(
                    consecutivo=consecutivo, 
                    lista_items=st.session_state['carrito'], 
                    solicitante=st.session_state['usuario_actual'], 
                    es_valido=True, 
                    motivo="Pedido reglamentario", 
                    ruta_pdf=ruta_pdf, 
                    justificacion=justificacion_global
                )

                bytes_pdf = b""
                if os.path.exists(ruta_pdf):
                    with open(ruta_pdf, "rb") as f_pdf:
                        bytes_pdf = f_pdf.read()

                st.session_state['ultimo_pedimento'] = {
                    "consecutivo": consecutivo,
                    "bytes_pdf": bytes_pdf,
                    "nombre_pdf": f"Pedimento_{consecutivo}.pdf"
                }
                st.session_state['carrito'] = []
                st.rerun()
