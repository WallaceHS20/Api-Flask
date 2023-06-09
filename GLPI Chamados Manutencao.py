from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from time import sleep
import os, sys
import xlsxwriter

#----------------------------------------------------------------------------------------------
# Pega o caminho do script
dir_path = ''
if getattr(sys, 'frozen', False):
	dir_path = os.path.dirname(sys.executable)
elif __file__:
    dir_path = os.path.dirname(__file__)
    
# Nome da planilha
workbookNome = "Extrair Dados.xlsx"
    

# remove mensagem de erro USB
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging']) 
# Instala o chrome aplicando as opções
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options) 

workbook = xlsxwriter.Workbook(dir_path + '\\' + workbookNome)
worksheet1 = workbook.add_worksheet()
celulaPadrao = workbook.add_format()
celulaMaior = workbook.add_format()
cell_format = workbook.add_format()
merge_format= workbook.add_format({'align': 'left'})
merge_format_center= workbook.add_format({'align': 'center'})

cell_format.set_text_wrap()
merge_format.set_text_wrap()
merge_format_center.set_border(2)
merge_format.set_border(1)
merge_format.set_align('top')
cell_format.set_border(1)
celulaPadrao.set_border(1)
celulaMaior.set_border(2)

driver.get('http://italia.uaga.local/glpi')

# encontrar os elementos
valida = 0

while(valida == 0):
	try:
		login = driver.find_element(By.ID,'login_name')
		senha = driver.find_element(By.ID,'login_password')
		botao = driver.find_element(By.CSS_SELECTOR, 'input[type=submit]')

		# enviar dados para os inputs
		sleep (1)
		user = input('Digite o usuário: ')
		password = input('Digite a senha: ')

		print('Aguarde...')

		login.send_keys(user)
		senha.send_keys(password)

		botao.click()

		driver.find_element(By.ID, 'ui-id-6').click()
		valida = valida + 1
	except:
		print('Login ou senha incorretos!')
		driver.find_element(By.XPATH, '//*[@id="bloc"]/div[2]/a').click()
		sleep(2)

sleep(5)

# contador = (driver.find_element(By.XPATH,'//*[@id="ui-tabs-3"]/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[5]/td[2]')).text
contador = (driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[4]/td[2]')).text
contador = int(contador)
# print(contador)
print(f'Chamados encontrados: {contador}')

confirma = 0
while (confirma < 1):
	de = eval(input('Digite o início do chamado: !!OBS: Digite 0 para pesquisar todos!! '))
	if (de < 0 or de > contador):
		print('!!Digite um valor maior do que 0!!')
	else:
		confirma = 1
if (de == 0):
	ate = 0
else:
	while (confirma > 0):
		ate = eval(input('Digite o fim do chamado: '))
		if (ate > contador or ate < de):
			print('!!Digite um valor menor ou igual ao total de chamados e que seja maior que o início do chamado!!')
		else:
			confirma = 0

if (de == 0 or ate ==0):
	de = 0
	ate = 0
	contador = contador
	
else:
	contador = ate

# try:
# 	confirma = 0
# 	while (confirma < 1):
# 		de = eval(input('Digite o início do chamado !!OBS: Digite 0 para pesquisar todos!!: '))
# 		if (de < 0):
# 			print('!!Digite um valor maior do que 0!!: ')
# 		elif (de > contador):
# 			print(f'!!Digite um valor menor do que {contador}!!: ')
# 		else:
# 			confirma = 1
# 	if (de == 0):
# 		ate = 0
# 	else:
# 		while (confirma > 0):
# 			ate = eval(input('Digite o fim do chamado: \n'))

# 			if (ate > contador or ate < de):
# 				print('!!Digite um valor menor ou igual ao total de chamados e que seja maior que o início do chamado!!\n')
# 			else:
# 				confirma = 0
# 	# if (ate ==0):
# 	# 	de = 0
# 	# 	ate = 0
# 	# 	contador = contador
# 	# else:
# 	# 	contador = ate

# except Exception:
#     print(Exception)

print('Aguarde...')

try:
    # Entra nos chamados atribuídos
	driver.find_element(By.XPATH,'//*[@id="ui-tabs-3"]/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[4]/td[1]/a').click()
except Exception:
    print(Exception)

col0 = 0
col1 = 1
col2 = 2

lin0 = 0
lin1 = 1
lin2 = 2
lin3 = 3
lin4 = 4
lin5 = 5
lin6 = 6

pagina = 0

def fazerConsulta():
	global col0 
	global col1 
	global col2 
		
	global lin0 
	global lin1 
	global lin2 
	global lin3 
	global lin4 
	global lin5 
	global lin6	

	global confirma 
	global pagina


	#--------------------------------------------------Antes de entrar no chamado-------------------------

	retorno = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[2]')).text
	retorno = retorno.replace(' ','')

	worksheet1.merge_range(lin0,col0,lin0,col2,'CONTROLE DE CHAMADOS',merge_format_center)

	worksheet1.write(lin1,col0,'ID: \n'+retorno)

	categoria = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[12]')).text

	worksheet1.write(lin1,col1,'Categoria:\n '+categoria,cell_format)

	data = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[6]')).text

	worksheet1.write(lin1,col2,'Data: \n'+data[0:10],cell_format)

	titulo = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[3]')).text

	worksheet1.merge_range(lin2,col0,lin2,col2,'Título: '+titulo,merge_format)

	solicitante = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[10]')).text

	worksheet1.merge_range(lin4,col0,lin4,col1,'Solicitante: \n'+solicitante[0:20],merge_format)

	solucao = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[8]')).text

	worksheet1.merge_range(lin5,col0,lin5,col1,'Data para solução: \n'+solucao[0:10],merge_format)

	prioridade = (driver.find_element(By.XPATH,'//*[@id="massformTicket"]/div/table/tbody/tr['+str(pagina)+']/td[7]')).text

	worksheet1.write(lin5,col2,'Prioridade: \n'+prioridade,cell_format)

	worksheet1.merge_range(lin6,col0,lin6,col2,'OBS: ',merge_format)

    #--------------------------------------------------Depoois de entrar no chamado-------------------------

	driver.find_element(By.XPATH,'//*[@id="Ticket'+retorno+'"]').click() #Click para entrar no chamado

	sleep(20)
        
	#descricao = (driver.find_element_by_xpath('//*[@id="mainformtable4"]/tbody/tr[2]/td')).text
	iframe = driver.find_element(By.XPATH, '//*[@id="mceu_24"]/iframe')
	driver.switch_to.frame(iframe)
	descricao = (driver.find_element(By.XPATH, '//*[@id="tinymce"]')).text
	
	# print(descricao)

	worksheet1.merge_range(lin3,col0,lin3,col2,'Descrição: '+descricao.rstrip('\n'),merge_format)
    
	driver.switch_to.default_content()

	localizacao = (driver.find_element(By.XPATH,'//*[@id="mainformtable2"]/tbody/tr[4]/td[2]/span/span[1]/span[1]/span')).text

	worksheet1.write(lin4,col2,'Localização: \n'+localizacao,cell_format)

	#----------------------Incremento para escrever nas posiçoes corretas das celulas-----------------------
	col0 = col0 + 4
	col1 = col1 + 4
	col2 = col2 + 4

	if(col0 == 12):
		col0 = 0
		col1 = 1
		col2 = 2
		lin0 = lin0 + 8
		lin1 = lin1 + 8
		lin2 = lin2 + 8
		lin3 = lin3 + 8
		lin4 = lin4 + 8
		lin5 = lin5 + 8
		lin6 = lin6 + 8

	driver.find_element(By.XPATH,'//*[@id="c_ssmenu2"]/ul/li[2]/a').click() #Click para sair do chamado

#--------------------------------Looping para chamar a função-----------------------------------------
for i in range(contador):

	pagina = pagina+1

	if (pagina < de): 

		contador = contador

	elif(pagina >= de):

		fazerConsulta()


#--------------------------------------Largura das Colunas do Excel----------------------------------------

worksheet1.set_column('A:A',10,celulaPadrao)
worksheet1.set_column('B:B',12,celulaPadrao)
worksheet1.set_column('C:C',15.5,celulaPadrao)

worksheet1.set_column('D:D',0.1,celulaMaior)

worksheet1.set_column('E:E',10,celulaPadrao)
worksheet1.set_column('F:F',12,celulaPadrao)
worksheet1.set_column('G:G',15.5,celulaPadrao)

worksheet1.set_column('H:H',0.1,celulaMaior)

worksheet1.set_column('I:I',10,celulaPadrao)
worksheet1.set_column('J:J',12,celulaPadrao)
worksheet1.set_column('K:K',15.5,celulaPadrao)

#--------------------------------------Altura das linhas do Excel----------------------------------------

worksheet1.set_row(0,13)
worksheet1.set_row(1,32)
worksheet1.set_row(2,34)
worksheet1.set_row(3,75)
worksheet1.set_row(4,32)
worksheet1.set_row(5,31)
worksheet1.set_row(6,30)

worksheet1.set_row(7,1,celulaMaior)

worksheet1.set_row(8,13)
worksheet1.set_row(9,32)
worksheet1.set_row(10,34)
worksheet1.set_row(11,75)
worksheet1.set_row(12,32)
worksheet1.set_row(13,31)
worksheet1.set_row(14,30)

worksheet1.set_row(15,1,celulaMaior)

worksheet1.set_row(16,13)
worksheet1.set_row(17,32)
worksheet1.set_row(18,34)
worksheet1.set_row(19,75)
worksheet1.set_row(20,32)
worksheet1.set_row(21,31)
worksheet1.set_row(22,30)

worksheet1.set_row(23,1,celulaMaior)

worksheet1.set_row(24,13)
worksheet1.set_row(25,32)
worksheet1.set_row(26,34)
worksheet1.set_row(27,75)
worksheet1.set_row(28,32)
worksheet1.set_row(29,31)
worksheet1.set_row(30,30)

worksheet1.set_row(31,1,celulaMaior)

worksheet1.set_row(31,13)
worksheet1.set_row(32,32)
worksheet1.set_row(33,34)
worksheet1.set_row(34,75)
worksheet1.set_row(35,32)
worksheet1.set_row(36,31)
worksheet1.set_row(37,30)

worksheet1.set_row(38,1,celulaMaior)

worksheet1.set_row(39,13)
worksheet1.set_row(40,32)
worksheet1.set_row(41,34)
worksheet1.set_row(42,75)
worksheet1.set_row(43,32)
worksheet1.set_row(44,31)
worksheet1.set_row(45,30)

worksheet1.set_row(46,1,celulaMaior)

worksheet1.set_row(47,13)
worksheet1.set_row(48,32)
worksheet1.set_row(49,34)
worksheet1.set_row(50,75)
worksheet1.set_row(51,32)
worksheet1.set_row(52,31)
worksheet1.set_row(53,30)

worksheet1.set_row(54,1,celulaMaior)

worksheet1.set_row(55,13)
worksheet1.set_row(56,32)
worksheet1.set_row(57,34)
worksheet1.set_row(58,75)
worksheet1.set_row(59,32)
worksheet1.set_row(60,31)
worksheet1.set_row(61,30)

worksheet1.set_row(62,1,celulaMaior)

# Fecha a planilha
workbook.close()
# Fecha o navegador
driver.quit()

print('Planilha gerada!')

sleep(2)
#--------------------------------------------------------------------------------------------