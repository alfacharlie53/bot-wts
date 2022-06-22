from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


import pandas as pd

class Bot(DesktopBot):

    def action(self, execution=None):

        # Opens the Whatsapp website.
        self.browse("http://web.whatsapp.com")

        # importar a base de dados
        tabela = pd.read_excel("C:/Users/invac/OneDrive/Área de Trabalho/bot whatapp envia msg/04-21 - Automatizar Envio de Mensagem do Whatsapp/botyoutube/botyoutube/Contatos.xlsx")
        print(tabela)
        # para cada linha da base de dados
        for linha in tabela.index:

            contato = tabela.loc[linha, "Contato"]
            msg = tabela.loc[linha, "Msg"]
            arquivo = tabela.loc[linha, "Arquivo"]
            

            if not self.find( "lupa", matching=0.97, waiting_time=60000):
                self.not_found("lupa")
            self.click()

            self.type_keys_with_interval(100, contato)
            self.enter()

            if pd.isna(arquivo): # tem arquivo
                self.paste(msg)
                self.enter()
            else:
                self.paste(msg)
                self.enter()
                if not self.find( "anexar", matching=0.97, waiting_time=10000):
                    self.not_found("anexar")
                self.click()
                if not self.find( "documento", matching=0.97, waiting_time=10000):
                    self.not_found("documento")
                self.click()
                if not self.find( "nome", matching=0.97, waiting_time=10000):
                    self.not_found("nome")
                self.paste(arquivo)
                self.enter()
                if not self.find( "enviar", matching=0.97, waiting_time=10000):
                    self.not_found("enviar")
                self.click()
                

            if self.find( "seta", matching=0.97, waiting_time=10000):
                self.click()

            self.wait(2000)
        
        
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()












