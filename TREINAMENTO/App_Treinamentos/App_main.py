# -*- coding: utf-8 -*-
import os
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox
import pandas as pd
import xlrd
import zipfile

from gui_main import Ui_ControleTreinamento

class FrontEnd(QDialog):
    def __init__(self):
        super(FrontEnd, self).__init__()
        
    #GUI - interface
        self.gui = Ui_ControleTreinamento()
        self.gui.setupUi(self)

        self.matrizpath = r"\\flsprd04\\grupos\\Treinamento Técnico\\2. MATRIZ DE TREINAMENTO\\1. MATRIZES VIGENTES"

        self.gui.btn_cliqueAqui.clicked.connect(self.selecionarDiretorio)

        self.docFile = None

    def selecionarDiretorio(self):
        try:
            pop = self.LerTexto()
            if pop:
                options = QFileDialog.Options()
                options |= QFileDialog.ShowDirsOnly
                folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", options=options)
                
                self.docFile = os.path.join(folder_path, f'dados_levantamento_{pop}.xlsx')

                self.executar_levantamento(pop)
            else:
                QMessageBox.critical(self, 'Erro!', 'POP inválido!')
        except:
            QMessageBox.critical(self, 'Erro!', 'Arquivo não gerado')

    def LerTexto(self):
        try:
            pop = self.gui.ln_pop.text().upper()
            return pop
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}")  
    
    def LimparText(self):
        try:
            self.gui.ln_pop.clear()
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}")  

    def executar_levantamento(self, levatamento_pop):
        try:
            caminhos = self.obter_caminhos_arquivos()
            # print(levatamento_pop, caminhos)
            if caminhos:
                dados = self.processar_arquivos(caminhos, levatamento_pop)
                if dados:
                    self.salvar_dados_excel(dados)
                    QMessageBox.information(self, 'Arquivo Gerado', 'Arquivo gerado com sucesso!')
                    self.LimparText()
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}")  

    def obter_caminhos_arquivos(self):
        try:
            path = r'\\flsprd04\\grupos\\Treinamento Técnico\\2. MATRIZ DE TREINAMENTO\\1. MATRIZES VIGENTES'
            diretorios = os.listdir(path=path)
            caminhos = []

            for pasta in diretorios:
                pasta_matriz = os.path.join(path, pasta)
                try:
                    for arquivo in os.listdir(pasta_matriz):
                        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls') or arquivo.endswith('.xlsb'):
                            caminhos.append(os.path.join(pasta_matriz, arquivo))
                except NotADirectoryError:
                    pass
            return caminhos
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}") 

    def processar_arquivos(self, caminhos, levatamento_pop):
        try:
            dados = {'Matrícula': [], 'Nome': [], 'Cargo': [], 'CDC': [], 'POP': [], 'Versão': [], 'Matriz': []}

            for caminho in caminhos:
                try:
                    dataframe_matriz = pd.read_excel(caminho, sheet_name='MATRIZ')
                    self.gravar_info(dataframe_matriz, os.path.basename(caminho), dados, levatamento_pop)
                except (PermissionError, xlrd.biffh.XLRDError, ValueError, zipfile.BadZipFile):
                    pass
            return dados
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}") 
    
    def gravar_info(self, dataframe, arquivo, dados, levatamento_pop):
        try:
            for i in range(4, len(dataframe.loc[6]), 3):
                for j in range(len(dataframe)):
                    if type(dataframe.loc[j][0]) == type(1):
                        if levatamento_pop == dataframe.loc[j - 2][i]:
                            try:
                                if dataframe.loc[j][i] == 'X':
                                    dados['Matrícula'].append(dataframe.loc[j][0])
                                    dados['Nome'].append(dataframe.loc[j][1])
                                    dados['Cargo'].append(dataframe.loc[j][3])
                                    dados['CDC'].append(dataframe.loc[j][2])
                                    dados['POP'].append(dataframe.loc[j - 2][i])
                                    dados['Versão'].append(dataframe.loc[j - 3][i])
                                    dados['Matriz'].append(arquivo)
                            except IndexError:
                                pass
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}") 
    
    def salvar_dados_excel(self, dados):
        try:
            caminho = self.docFile
            tabela_pops = pd.DataFrame(dados)
            tabela_pops.to_excel(caminho, index=False)
        except Exception as e:
            print(f"Ocorreu uma exceção: {e}") 

    def closeEvent(self, event):
        event.accept()        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    gui = FrontEnd()
    gui.show() 
    sys.exit(app.exec_())
