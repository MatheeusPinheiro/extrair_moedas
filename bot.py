# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

#atualizar o chromedriver automaticamente
from webdriver_manager.chrome import ChromeDriverManager

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

import pandas as pd
import os

#Email
import win32com.client as win32


class Bot(WebBot): 


    def search_coin(self, moeda):
        self.find_element('APjFqb', By.ID).clear()
        self.find_element('APjFqb', By.ID).send_keys(moeda)
        self.enter()

    def extract_coin(self):
        coin = self.find_element('pclqee', By.CLASS_NAME).text
        return coin
    
    def send_email(self, email_to, file_path):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email_to
        mail.Subject = 'Valores das Moedas'
        mail.Body = 'Segue em anexo os valores das moedas.'

        # Anexa o arquivo
        mail.Attachments.Add(file_path)
        # Envia o e-mail
        mail.Send()

        print("E-mail enviado com sucesso!")
    


    def action(self, execution=None):
        # Runner passes the server url, the id of the task being executed,
        # the access token and the parameters that this task receives (when applicable).
        maestro = BotMaestroSDK.from_sys_args()
        ## Fetch the BotExecution with details from the task, including parameters
        execution = maestro.get_execution()

        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")


        # Configure whether or not to run on headless mode
        self.headless = False

        # Uncomment to change the default Browser to Firefox
        self.browser = Browser.CHROME

        # Uncomment to set the WebDriver path
        self.driver_path = ChromeDriverManager().install()

        # Opens the BotCity website.
        self.browse("https://www.google.com/")

        #Maximize Window
        self.maximize_window()

        # Implement here your logic...
        
        try:
            #Alerta para inicio da automação
            maestro.alert(
                task_id=execution.task_id,
                title="Extrair_Moedas - Inicio",
                message="Estamos iniciando o processo",
                alert_type=AlertType.INFO
            )

            #Parametros vindo do orquestrador
            coin1 = execution.parameters.get("moeda")
            coin2 = execution.parameters.get("moeda2")
            
            #Caminho da pasta
            pasta = os.getcwd()
            
            #Procuro a moeda dolar
            self.search_coin(coin1)
            self.wait(1000)
            dolar =  self.extract_coin()

            #Procuro a moeda euro
            self.wait(1000)
            self.search_coin(coin2)
            self.wait(1000)
            euro =  self.extract_coin()

            #Criando um dataframe no pandas
            data = {'Moedas': [coin1, coin2], 'Valor': [dolar, euro]}
            df = pd.DataFrame(data)

            #Salvando os dados na planilha do excel
            df.to_excel('moedas.xlsx',index=False)
            
            #Juntar o caminho da pasta com o arquivo excel
            file_xlsx = os.path.join(pasta, 'moedas.xlsx')

            #Aguarda 2s
            self.wait(3000)
            
            #Enviando e-mail
            self.send_email('matheuspinheiro0597@gmail.com', file_xlsx)
            
 
        except Exception as ex:
            #Print da tela se que deu erro
            self.save_screenshot('erro.png')
            
            #Mensagem de erro
            maestro.error(
                task_id=execution.task_id,
                exception=ex,
                screenshot="erro.png"
            )
            
            #Status da tarefa
            status = AutomationTaskFinishStatus.FAILED
            message = "Tarefa extrair moedas finalizada com falha"

        finally:

            # Wait 3 seconds before closing
            self.wait(3000)

            #Encerrar o Browser
            self.stop_browser()

            # Uncomment to mark this task as finished on BotMaestro
            maestro.finish_task(
                task_id=execution.task_id,
                status=AutomationTaskFinishStatus.SUCCESS,
                message="Tarefa finalizada OK."
            )


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
