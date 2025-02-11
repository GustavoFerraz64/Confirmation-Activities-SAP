import time
import sys
import win32com.client
import pandas as pd
import traceback
import time

#Define o caminho onde o arquivo extraído da transação ZP030 ficará armazenado.
PATH_ZP030 = r'C:\Temp' 

#Define o caminho onde está localizado o arquivo "Obras CN33 (ZP030 e CN47N).xlsx"
PATH_OBRAS = r'\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\CN33\Macro\Obras (Atividade 70).xlsx'

#Define o caminho onde será salvo o report do resultado.
PATH_RESULTADO = r'\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\CN33\Macro\Obras (Atividade 70)_Resultado.xlsx' 

"""
Script de automação do processo de confirmação da atividade 70.

Ao executar o script, o robô irá conferir se as obras informadas no arquivo excel estão com pendências na transação ZP030. As obras que não estiverem pendentes, terão
a sua atividade 70 confirmada na transação CN47N. Já as obras com pendências na transação ZP030, não terão a atividade 70 confirmadas.

Ao final do script, será gerado um arquivo chamado "Obras (Atividade 70)_Resultado.xlsx" que irá informar sobre o status de cada obra em relação às transações ZP030 e CN47N.
O novo arquivo será gerado na mesma pasta do arquivo preenchido.

Instruções: 
    -Colocar as obras (primeira coluna) e seu respectivo elemento PEP (segunda coluna) no arquivo "Obras (Atividade 70).xlsx", localizado em:
    \\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\CN33\Macro

    -Após preencher a coluna Obra e a coluna Elemento PEP no arquivo "Obras (Atividade 70).xlsx", salve e feche o arquivo Excel.

    -Antes de executar o código, certifique-se de que esteja logado no SAP.

@autor: Gustavo Nunes Ferraz
@data : 24/04/2024
@departamento: DPCP
@modificado: 06/08/2024

Histórico de modificações:

-20/06/2024: Obras com status "Stat.mat.espec.cent." igual a 01 serão passadas para a confirmação na CN47N
-06/08/2024: Obras com centro igual a 5001 serão passadas para a confirmação na CN47N

"""

class ATV70:
   
    """Classe destinada a automatizar o processo de confirmação da atividade 70.
    
    Funções principais:
    -------------
    
    - Verificar se o SAP está aberto
    - Verificar se há um usuário logado
    - Fazer as verificações necessárias para determinar se as obras estão com alguma pendência na ZP030
    - Confirmar a atividade 70 na CN47N, caso a obra não possua pendência
    
    """

    def __init__(self):
        """Pega a aplicação COM do SAP e a primeira sessão para uso no script
           Testa se o SAP está aberto e se há um usuário logado
           
           """
        try: #Garante que o SAP está aberto, caso contrário encerra o programa
        #Pega a aplicação COM do SAP e a primeira sessão para uso no script
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            self.application = self.SapGuiAuto.GetScriptingEngine
            self.connection = self.application.Children(0)
            self.sessions = self.connection.Children
            self.session = self.sessions[0]
        except:
            print("SAP não está aberto, o programa será finalizado")
            time.sleep(2)
            sys.exit(0)
        #Testa se há um usuário logado no SAP, caso não haja, encerra o programa
        self.usuario = self.session.Info.User
        if self.usuario == '':
            print("SAP não está logado, o programa será finalizado")
            time.sleep(2)
            sys.exit(0)

    def zp030(self):
        """Acessa a ZP030 com as obras a serem rodadas para determinar se ainda existe alguma pendência de material

        Returns:
            Bool: flag que é True se todas as obras não tiverem problema e False caso contrário
        """
        
        session = self.session
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n zp030"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/btn%_S_ZZNRO_%_APP_%-VALU_PUSH").press() #Seleção múltipla da obra
        session.findById("wnd[1]/tbar[0]/btn[16]").press() #Limpa todas as entradas
        session.findById("wnd[1]/tbar[0]/btn[24]").press() #Realiza o upload to clipboard
        session.findById("wnd[1]/tbar[0]/btn[8]").press() #Aplica o filtro das obras
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Roda a transação
        if session.findById("wnd[0]/sbar").text == 'Não existem dados': #Testa se houve problema com alguma obra
            return True #Se não houve, retorna True
        else: #Extração da ZP030 com os problemas
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select() #Seleciona extração de planilha eletrônica
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = PATH_ZP030 #Diretório
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zp030.XLSX" #Nome do arquivo
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 1
            session.findById("wnd[1]/tbar[0]/btn[11]").press() #Substituir

            #Verificar se o arquivo foi salvo corretamente            
            if session.findById("wnd[0]/sbar").text == '':
                print("A extração da planilha na transação ZP030 não foi salva corretamente, certifique-se de que a planilha 'zp030.xlsx' esteja fechada durante a execução do script")
                print("O programa será finalizado")
                sys.exit(0)    
            return False

    def cn47n_conf(self, obras_cn47n, df_obras):
        """Acessa a CN47N com os PEPs para rodar a atividade 60 das obras que não tiveram nenhum problema de verificação

        Returns:
            Pandas DataFrame: dataframe com o retorno referente à cada obra (status e mensagem SAP).
        """
        
        session = self.session 
        session.findById("wnd[0]/tbar[0]/okcd").text = r"/ncn47n"
        session.findById("wnd[0]").sendVKey(0)
        
        #Verifica se a tela de seleção de Perfil-BD aparece
        try: #Se aparecer, seleciona a padrão
            session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "000000000001"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: #Se não aparecer, troca o Perfil-BD para o padrão dentro da transação
            session.findById("wnd[0]/tbar[1]/btn[28]").press() #Botão de troca Perfil-BD
            session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "000000000001"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
        session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").text = ''
        session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press() #Seleção múltipla de PEPs
        session.findById("wnd[1]/tbar[0]/btn[16]").press() #Limpa todas as entradas
        session.findById("wnd[1]/tbar[0]/btn[24]").press()#Realiza o upload to clipboard
        session.findById("wnd[1]/tbar[0]/btn[8]").press() #Filtra os PEPs
        session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = "0070" #Coloca atividade 0060 no low
        session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/ELEMENTPEP" #Selelciona o layout
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Roda a transação
        
        #Inicia a confirmação para cada obra na CN47N
        for index, row in obras_cn47n.iterrows():
            session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").setCurrentCell(index,"")
            session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectedRows = index #Seleciona a primeira linha
            session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").pressToolbarContextButton("CONF") #Clica no botão de confirmação
            session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem("CREACONF") #Cria uma nova confirmação

            #Tenta fazer a confirmação, no caso de não aparecer uma janela Pop Up
            try:
                session.findById("wnd[0]/usr/chkAFRUD-AUERU").selected = True #Seleciona a flag de Conf.final
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                df_obras.loc[df_obras['Obra'] == row['Obra'], 'SAP CN47N'] = session.findById("wnd[0]/sbar").text #Report para o relatório
                df_obras.loc[df_obras['Obra'] == row['Obra'], 'CN47N (Atividade 70)'] = 'OK' #Report para o relatório
                continue
            except:

                #Tenta fazer a confirmação, no caso de aparecer uma janela Pop Up
                try:
                    session.findById("wnd[1]/usr/chkAFRUD-AUERU").selected = True #Seleciona a flag de Conf.final
                    session.findById("wnd[1]/tbar[0]/btn[11]").press() #Salva
                    session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    df_obras.loc[df_obras['Obra'] == row['Obra'], 'SAP CN47N'] = session.findById("wnd[0]/sbar").text #Report para o relatório
                    df_obras.loc[df_obras['Obra'] == row['Obra'], 'CN47N (Atividade 70)'] = 'OK' #Report para o relatório
                    continue
                except:

                    #Caso aparecer uma mensagem de erro, tenta fechar a mensagem e voltar para confirmar outras obras
                    try:
                        session.findById("wnd[1]/usr/lblERROR")
                        df_obras.loc[df_obras['Obra'] == row['Obra'], 'SAP CN47N'] = 'Erro' #Report para o relatório
                        df_obras.loc[df_obras['Obra'] == row['Obra'], 'CN47N (Atividade 70)'] = 'Erro' #Report para o relatório
                        session.findById("wnd[1]").close()
                        session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        continue

                    except:

                        #Caso aparecer um erro específico, tenta fechar a mensagem e voltar para confirmar outras obras
                        if session.findById("wnd[0]/sbar").text == 'Tipo de capacidade e nº partição têm que ser atualizados juntamente':
                            df_obras.loc[df_obras['Obra'] == row['Obra'], 'SAP CN47N'] = session.findById("wnd[0]/sbar").text #Report para o relatório
                            df_obras.loc[df_obras['Obra'] == row['Obra'], 'CN47N (Atividade 70)'] = 'Erro' #Report para o relatório
                            session.findById("wnd[0]/tbar[0]/btn[3]").press()
                            continue

                        else:
                            if session.findById("wnd[1]/usr/chkAFRUD-AUERU", False): #Teste para verificar se a atividade já foi rodada
                                session.findById("wnd[1]/usr/chkAFRUD-AUERU").selected = True #Seleciona a flag de Conf.final
                                session.findById("wnd[1]/tbar[0]/btn[11]").press() #Salva
                                session.findById("wnd[2]").sendVKey(0) #Dá ok no aviso
                            else: #Se já foi, vai aparecer um aviso de que isso aconteceu
                                session.findById("wnd[1]/usr/btnOPTION1").press() #Clica no botão "Não", para não confirmar novamente a atividade
                                session.findById("wnd[0]/tbar[0]/btn[3]").press() #Volta para a tela das confirmações

                            #Informa no arquivo excel que a  atividade 70 foi confirmada, junto com as mensagens emitidas pelo SAP.
                            df_obras.loc[df_obras['Obra'] == row['Obra'], 'SAP CN47N'] = session.findById("wnd[0]/sbar").text
                            df_obras.loc[df_obras['Obra'] == row['Obra'], 'CN47N (Atividade 70)'] = 'OK'

def main():
    try:
        print("Iniciando o script...")

        #Cria a instância da classe que executa a atividade 70
        print("Fazendo verificações com o SAP...")
        atv70 = ATV70()
        print('Concluído!')

        #Cria um dataframe com o arquivo excel contendo a Obra e seu respectivo Elemento PEP, que serão executados nas transações ZP030 e CN47N
        print("Lendo o arquivo Obras '(Atividade 70).xlsx'...")
        df_obras = pd.read_excel(PATH_OBRAS)

        #Adiciona as colunas para informar o status de cada obra, após a execução do script
        df_obras['ZP030'] = 'Pendente'
        df_obras['CN47N (Atividade 70)'] = '-'
        df_obras['SAP CN47N'] = '-'

        #Copia todas as obras na coluna obras, do arquivo excel lido
        df_obras['Obra'].to_clipboard(index = False, header= False)

        print("Executando a transação ZP030...")

        #Se há alguma pendência na ZP030, é executado o código abaixo
        if not atv70.zp030():

            print("Informando as pendências encontradas na transação ZP030...")

            #O arquivo extraído da transação ZP030 é armazenado na pasta \Temp
            df_zp030 = pd.read_excel(PATH_ZP030 + r'\zp030.xlsx') 
            df_zp030 = df_zp030.loc[df_zp030['Stat.mat.espec.cent.'] != 1] #Obras que tem Stat. mat igual a 1 terão suas atividades confirmadas
            

            #Se algum Elemento PEP da planilha original lida não estiver no arquivo gerado pela transação ZP030, a coluna ZP030 do arquivo original é marcado como "OK"
            df_obras.loc[~df_obras['Elemento PEP'].isin(df_zp030['Elemento PEP']), 'ZP030'] = 'OK'  

        else:
            #Caso não há nenhum resultado da transação ZP030, significa que nenhuma obra está pendente.
            print("Transação ZP030 concluída, nenhuma obra com pendência foi encontrada.")
            df_obras["ZP030"] = 'OK'
        
        #Novo dataframe criado, filtrando somente as obras que não estiverem com ZP030 marcado como erro 
        obras_cn47n = df_obras.loc[df_obras['ZP030'] == 'OK']

        #As obras do novo dataframe criado são copiadas
        obras_cn47n['Elemento PEP'].to_clipboard(index = False, header = False)
        obras_cn47n.reset_index(drop=True, inplace=True)

        #Verifica se caso todas as obras informadas no arquivo excel estão com pendência na ZP030, caso positivo, o script será finalizado
        if df_obras.loc[df_obras['ZP030'] != 'Pendente'].empty:
            
            print("Todas as obras informadas no arquivo excel estão com pendências na ZP030, a transação CN47N não será executada.")

            df_obras.to_excel(PATH_RESULTADO, index=False)

        #Caso contrário, é acessado a transação CN47N para confirmar as obras que não apresentaram pendências nas ZP030    
        else:
            print("Confirmando a atividade 70 na transação CN47N...")
            atv70.cn47n_conf(obras_cn47n, df_obras)
            print("Salvando o arquivo Excel com as informações...")
            df_obras.to_excel(PATH_RESULTADO, index=False)

        print("Processo finalizado.")
        time.sleep(2)
        
    except Exception as e:
        print(e)
        traceback.print_exc()

if __name__ == '__main__':
    main()