import streamlit as st
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE # Importado para formatar a lista de destinatários

# --- Configurações SMTP (Lidas do st.secrets) ---
try:
    SMTP_SERVER = st.secrets["smtp"]["servidor"]
    SMTP_PORT = st.secrets["smtp"]["porta"]
    REMETENTE_PADRAO = st.secrets["smtp"]["email_remetente"]
    SENHA_APP = st.secrets["smtp"]["senha_app"]
    
    # LISTA DE DESTINATÁRIOS ATUALIZADA AQUI:
    DESTINATARIOS_PADRAO = [
        "iasmin.fernandes@lcmconstrucao.com.br", 
        "grp.contabil@lcmconstrucao.com.br", 
        "maria.eliza@lcmconstrucao.com.br"
    ]
except KeyError:
    st.error("ERRO: As credenciais SMTP não foram configuradas corretamente em .streamlit/secrets.toml")
    SMTP_SERVER, SMTP_PORT, REMETENTE_PADRAO, SENHA_APP = "", 587, "", ""
    DESTINATARIOS_PADRAO = []


# --- Função de Envio de E-mail (SMTP) ---

def enviar_email_smtp(remetente, senha, destinatarios, assunto, corpo_texto, corpo_html):
    """Envia um e-mail através de um servidor SMTP, com formato texto e HTML, para múltiplos destinatários."""
    
    try:
        # Cria a mensagem como 'alternative'
        msg = MIMEMultipart('alternative')
        msg['From'] = remetente
        
        # Junta a lista de destinatários em uma string para o campo 'To'
        msg['To'] = COMMASPACE.join(destinatarios) 
        msg['Subject'] = assunto
        
        # Adiciona o corpo em texto simples (fallback)
        msg.attach(MIMEText(corpo_texto, 'plain'))
        
        # Adiciona o corpo em HTML
        msg.attach(MIMEText(corpo_html, 'html'))

        # Conecta ao servidor e envia o e-mail
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(remetente, senha)
        text = msg.as_string()
        
        # O sendmail aceita a lista de destinatários
        server.sendmail(remetente, destinatarios, text) 
        server.quit()
        
        return True, "E-mail enviado com sucesso para todos os destinatários!"
        
    except smtplib.SMTPAuthenticationError:
        return False, "Falha na autenticação SMTP. Verifique a Senha de Aplicativo (App Password)."
    except Exception as e:
        return False, f"Erro ao enviar o e-mail: {e}"

# --- Função Principal de Processamento (MODIFICADA para Empresa_nf) ---

@st.cache_data
def processar_planilhas(uploaded_prefeitura, uploaded_uau):
    
    # 1. Leitura e Filtragem do 'prefeitura.xlsx'
    try:
        df_prefeitura = pd.read_excel(uploaded_prefeitura)
        
        # Colunas esperadas da Prefeitura: Número, Situação Documento, Data Emissão
        df_cancelados = df_prefeitura[df_prefeitura['Situação Documento'] == 'Cancelado'][
            ['Número', 'Situação Documento', 'Data Emissão'] 
        ].copy() 
        
        if df_cancelados.empty:
            return None, None, "Nenhum documento cancelado foi encontrado na planilha da Prefeitura."

    except KeyError as e:
        return None, None, f"ERRO: A coluna {e} ou outra coluna essencial não foi encontrada na planilha da Prefeitura. As colunas esperadas são: 'Número', 'Situação Documento' e 'Data Emissão'."
    except Exception as e:
        return None, None, f"ERRO ao ler a planilha da Prefeitura: {e}"


    # 2. Leitura e Preparação do 'uau.xlsx' (MODIFICADA para Empresa_nf)
    try:
        df_uau = pd.read_excel(uploaded_uau)
        
        # MODIFICAÇÃO: Incluindo 'Empresa_nf' na seleção de colunas do UAU
        df_uau_cols = df_uau[['NumNfAux_nf', 'Status_nf', 'Empresa_nf']].copy()
        df_uau_cols.rename(columns={'NumNfAux_nf': 'Número'}, inplace=True)
        
        # Mapeamento Status (0/1)
        status_map = {0: 'Normal', 1: 'Cancelado'}
        df_uau_cols['Status_uau'] = df_uau_cols['Status_nf'].fillna(-1).astype(int).map(status_map)
        
        # Selecionando as colunas finais do UAU
        df_uau_cols = df_uau_cols[['Número', 'Status_uau', 'Empresa_nf']]
        
    except KeyError as e:
        return None, None, f"ERRO: A coluna {e} ou outra coluna essencial não foi encontrada na planilha UAU. As colunas esperadas são: 'NumNfAux_nf', 'Status_nf' e 'Empresa_nf'."
    except Exception as e:
        return None, None, f"ERRO ao ler a planilha UAU: {e}"


    # 3. Junção (Merge) dos dados
    df_resultado = pd.merge(
        df_cancelados,
        df_uau_cols,
        on='Número',
        how='left'
    )
    
    # 4. Geração e Tratamento dos Resultados
    
    # 4.1. Coluna VERIFICADO (Existência em UAU)
    df_resultado['VERIFICADO'] = df_resultado['Status_uau'].notna()
    existencia_map = {True: 'ENCONTRADO', False: 'NÃO ENCONTRADO'}
    df_resultado['VERIFICADO'] = df_resultado['VERIFICADO'].map(existencia_map)

    # 4.2. Coluna Situação UAU e Empresa UAU (Tratamento de Não Encontrado)
    df_resultado['Status_uau'].fillna('Não Encontrado', inplace=True)
    df_resultado['Empresa_nf'].fillna('Não Encontrado', inplace=True) # Preenche valores NaN se a NF não for encontrada
    
    # 5. Formatação Final da Saída (MODIFICADA para Empresa_nf)
    df_final = df_resultado[[
        'Número', 
        'Situação Documento', 
        'Data Emissão', 
        'VERIFICADO', 
        'Status_uau',
        'Empresa_nf' # Adicionado aqui
    ]].copy()
    
    # Renomeação das Colunas (MODIFICADA)
    novos_nomes = {
        'Número': 'Número NF',
        'Situação Documento': 'Situação Prefeitura',
        'Data Emissão': 'Data Emissão NF', 
        'VERIFICADO': 'Existencia UAU',
        'Status_uau': 'Situação UAU',
        'Empresa_nf': 'Empresa UAU' # <-- Novo nome para a coluna
    }
    df_final.rename(columns=novos_nomes, inplace=True)
    
    # 6. Geração do Filtro de Inconsistência
    
    df_inconsistencia = df_final[
        (df_final['Existencia UAU'] == 'ENCONTRADO') & 
        (df_final['Situação UAU'] == 'Normal')
    ].copy()
    
    return df_final, df_inconsistencia, None

# --- Interface Streamlit (Mantida) ---

st.set_page_config(
    page_title="Validação de Documentos Cancelados",
    layout="wide"
)

st.title("🔎 Validação de Documentos Cancelados (Prefeitura vs. UAU)")
st.markdown("Carregue as duas planilhas Excel para verificar a existência e o status das notas fiscais canceladas da Prefeitura na base UAU.")

# Área de upload de arquivos
col1, col2 = st.columns(2)
with col1:
    uploaded_prefeitura = st.file_uploader(
        "📤 1. Planilha da Prefeitura", 
        type=['xlsx'],
        key="prefeitura_uploader"
    )
with col2:
    uploaded_uau = st.file_uploader(
        "📤 2. Planilha UAU", 
        type=['xlsx'],
        key="uau_uploader"
    )

st.divider()

if uploaded_prefeitura and uploaded_uau:
    
    df_final, df_inconsistencia, error_message = processar_planilhas(uploaded_prefeitura, uploaded_uau)
    
    if error_message:
        st.error(error_message)
    elif df_final is None or df_final.empty:
        st.warning("Nenhum documento cancelado foi encontrado para análise.")
    else:
        # --- EXIBIÇÃO DA INCONSISTÊNCIA ---
        st.header("⚠️ Inconsistências Detectadas")
        st.markdown("**Documentos Cancelados na Prefeitura, mas Ativos/Normais no UAU.**")
        
        if not df_inconsistencia.empty:
            st.error(f"Encontrados **{len(df_inconsistencia)}** documentos em estado de inconsistência!")
            # O st.dataframe exibirá as novas colunas
            st.dataframe(df_inconsistencia, use_container_width=True)
            
            # --- PREPARAÇÃO DO EMAIL (CORPO HTML) ---
            assunto = f"[Ação Necessária] Inconsistências de NF Canceladas ({len(df_inconsistencia)} documentos)"
            
            # 1. Corpo em TEXTO PURO (Fallback) - Inclui as novas colunas automaticamente
            corpo_texto = "Prezados(as),\n\nForam detectadas as seguintes inconsistências em notas fiscais que estão 'Canceladas' na Prefeitura, mas 'Normais' (ativas) no sistema UAU. Favor verificar:\n\n"
            corpo_texto += df_inconsistencia.to_string(index=False)
            corpo_texto += f"\n\nAtenciosamente,\nRelatório Automático (Enviado por {REMETENTE_PADRAO})\nFavor não responder este e-mail, pois ele é gerado automaticamente.\n Favor responder ao e-mail: elzimar.mota@lcmconstrucao.com.br"
            
            # 2. Corpo em HTML (Com a Tabela Formatada!) - Inclui as novas colunas automaticamente
            tabela_html = df_inconsistencia.to_html(index=False) 

            # Template HTML
            corpo_html = f"""\
            <html>
              <body>
                <p>Prezados(as),</p>
                <p>Foram detectadas as seguintes inconsistências em notas fiscais que estão 'Canceladas' na Prefeitura, mas 'Normais' (ativas) no sistema UAU. Favor verificar:</p>
                
                {tabela_html}
                
                <br><p>Atenciosamente,</p><br>
                
                <p>Favor não responder este e-mail, pois ele é gerado automaticamente por {REMETENTE_PADRAO}.</p>
                <p>Se necessário, favor responder ao e-mail: <a href="mailto:elzimar.mota@lcmconstrucao.com.br">elzimar.mota@lcmconstrucao.com.br</a></p>
              </body>
            </html>
            """
            
            # Botões
            col_inc_dl, col_inc_mail = st.columns(2)
            
            # Botão de download
            excel_buffer_inc = io.BytesIO()
            df_inconsistencia.to_excel(excel_buffer_inc, index=False, engine='openpyxl')
            excel_buffer_inc.seek(0)
            with col_inc_dl:
                st.download_button(
                    label="💾 Baixar Inconsistências em Excel",
                    data=excel_buffer_inc,
                    file_name="relatorio_inconsistencias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # Botão de Enviar E-mail
            with col_inc_mail:
                if st.button("📧 Enviar E-mail", use_container_width=True):
                    with st.spinner('Enviando e-mail...'):
                        success, message = enviar_email_smtp(
                            remetente=REMETENTE_PADRAO,
                            senha=SENHA_APP,
                            destinatarios=DESTINATARIOS_PADRAO, 
                            assunto=assunto,
                            corpo_texto=corpo_texto,
                            corpo_html=corpo_html 
                        )
                        
                        if success:
                            st.success(message)
                        else:
                            st.error(f"Falha ao enviar e-mail: {message}")

        else:
            st.success("✅ Nenhuma inconsistência (Cancelado/Normal) encontrada!")

        st.divider()
        
        # --- EXIBIÇÃO DO RESULTADO COMPLETO ---
        st.header("Tabela de Resultados Completos")
        # O st.dataframe exibirá as novas colunas
        st.dataframe(df_final, use_container_width=True)
        st.success(f"Análise completa para **{len(df_final)}** documentos cancelados.")
        
        # Botão de download do resultado completo
        excel_buffer_full = io.BytesIO()
        df_final.to_excel(excel_buffer_full, index=False, engine='openpyxl')
        excel_buffer_full.seek(0)
        st.download_button(
            label="💾 Baixar Tabela Completa em Excel",
            data=excel_buffer_full,
            file_name="relatorio_cancelados_verificados_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Por favor, carregue ambas as planilhas para iniciar a verificação.")