#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[62]:


#Importar bibliotecas

import pandas as pd
import win32com.client as win32
import pathlib


# In[63]:


#Importar Base de Dados

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas =  pd.read_csv(r'Bases de Dados\Lojas.csv', sep=';', encoding= 'latin1')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

display(emails)
display(lojas)
display(vendas)


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[64]:


#Inclusão da coluna Loja no DF Vendas

vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


# In[65]:


dic_lojas = {}
for loja in lojas['Loja']:
    dic_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
    


# In[66]:


dia_indicador = vendas['Data'].max()
print(dia_indicador)


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[89]:


#Identificar se a pasta já existe


caminho_aqv = pathlib.Path(r'Backup Arquivos Lojas')
aqv_pasta_bkp =  caminho_aqv.iterdir() 

lista_backup = [arquivo.name for arquivo in aqv_pasta_bkp]


for loja in dic_lojas:
   if loja not in lista_backup:
       nova_pasta = caminho_aqv / loja
       nova_pasta.mkdir()
   
   nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month,dia_indicador.day, loja)
   local_arquivo  = caminho_aqv / loja / nome_arquivo
   
   dic_lojas[loja].to_excel(local_arquivo)


# ### Passo 4 - Calcular o indicador para 1 loja

# In[68]:


#meta

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


# In[95]:


for loja in dic_lojas:


    vendas_loja = dic_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]


    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)

    #diversidade

    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    #print(qtde_produtos_ano)
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    #print(qtde_produtos_dia)

    #ticket medio ano

    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_ano = valor_venda['Valor Final'].mean()
    #print(ticket_ano)

    #ticket medio dia

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_dia = valor_venda_dia['Valor Final'].mean()
    #print(ticket_dia)
    
        
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = 'OnePage Dia {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja)

    if faturamento_dia > meta_faturamento_dia:
        cor_fat_dia= 'green'
    else:
        cor_fat_dia= 'red'
    if faturamento_ano > meta_faturamento_ano:
        cor_fat_ano= 'green'
    else:
        cor_fat_ano= 'red'
    if qtde_produtos_dia > meta_qtdeprodutos_dia:   
        cor_qtde_dia= 'green'
    else:
        cor_qtde_dia= 'red'
    if qtde_produtos_ano > meta_qtdeprodutos_ano:
        cor_qtde_ano= 'green'
    else:
        cor_qtde_ano= 'red'
    if ticket_dia > meta_ticketmedio_dia:
        cor_ticket_dia= 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_ano > meta_ticketmedio_ano:
        cor_ticket_ano= 'green'
    else:
        cor_ticket_ano= 'red'


    mail.HTMLBody = '''

    <p> Bom dia, {nome}</p>

    <p> O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja}</strong> foi de:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Cenário Dia</th>
        <th>Meta Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style= "text-align: center">R${faturamento_dia:.2f}</td>
        <td style= "text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style= "text-align: center"><font color = {cor_fat_dia}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style= "text-align: center">{qtde_produtos_dia:.}</td>
        <td style= "text-align: center">{meta_qtdeprodutos_dia:.}</td>
        <td style= "text-align: center"><font color = {cor_qtde_dia}>◙</font>◙</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style= "text-align: center">R${ticket_dia:.2f}</td>
        <td style= "text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style= "text-align: center"><font color = {cor_ticket_dia}>◙</font>◙</td>
      </tr>

    </table>

    <br>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Cenário Ano</th>
        <th>Meta Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style= "text-align: center">R${faturamento_ano:.2f}</td>
        <td style= "text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style= "text-align: center"><font color = {cor_fat_ano}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style= "text-align: center">{qtde_produtos_ano}</td>
        <td style= "text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style= "text-align: center"><font color = {cor_qtde_ano}>◙</font>◙</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style= "text-align: center">R${ticket_ano:.2f}</td>
        <td style= "text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style= "text-align: center"><font color = {cor_ticket_ano}>◙</font>◙</td>
      </tr>

    </table>


    <p> Segue em anexo a planilha com todos os detalhes.</p>

    <p> Att., Thomas </p>
    '''

    # Anexos (pode colocar quantos quiser):
    attachment  = pathlib.Path.cwd() / caminho_aqv / loja / '{}_{}_{}.xlsx.'.format(dia_indicador.month,dia_indicador.day, loja )
    mail.Attachments.Add(str(attachment))

    mail.Send()
    
    print('O E-mail da loja {} foi enviado'.format(loja))
    


# ### Passo 5 - Criar ranking para diretoria

# In[91]:


faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas = faturamento_lojas.sort_values(by='Valor Final', ascending=False)


nome_arquivo = '{}_{}_Ranking_Anual.xlsx'.format(dia_indicador.month,dia_indicador.day)
faturamento_lojas.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


vendas_dia = vendas.loc[vendas['Data']==dia_indicador , :]
faturamento_dias = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_dias = faturamento_dias.sort_values(by='Valor Final', ascending=False)


nome_arquivo = '{}_{}_Ranking_Diario.xlsx'.format(dia_indicador.month,dia_indicador.day)
faturamento_dias.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


display(faturamento_lojas)
display(faturamento_dias)


# ### Passo 6 - Enviar e-mail para diretoria

# In[94]:


outlook = win32.Dispatch('outlook.application')

nome = emails.loc[emails['Loja']=='Diretoria', 'Gerente'].values[0]
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
mail.Subject = 'Rankin do Dia {}/{}'.format(dia_indicador.day, dia_indicador.month)


mail.Body = f'''

Prezados, bom dia.

Melhor Loja do dia em faturamento: Loja {faturamento_dias.index[0]} com faturamento de R${faturamento_dias.iloc[0, 0]:.2f}
Pior Loja do dia em faturamento: Loja{faturamento_dias.index[-1]} com faturamento de R${faturamento_dias.iloc[-1, 0]:.2f}

Melhor Loja do ano em faturamento: Loja {faturamento_lojas.index[0]} com faturamento de R${faturamento_lojas.iloc[0, 0]:.2f}
Pior Loja do ano em faturamento: Loja {faturamento_lojas.index[-1]} com faturamento de R${faturamento_lojas.iloc[-1, 0]:.2f}

Segue em anexo o ranking anual e diário de todas as lojas

Qualquer dúvida, estou à disposição.

Atte. Thomas

'''

# Anexos (pode colocar quantos quiser):
attachment  = pathlib.Path.cwd() / caminho_aqv /'{}_{}_{}.xlsx.'.format(dia_indicador.month,dia_indicador.day, 'Ranking_Anual')
mail.Attachments.Add(str(attachment))

attachment  = pathlib.Path.cwd() / caminho_aqv/ '{}_{}_{}.xlsx.'.format(dia_indicador.month,dia_indicador.day, 'Ranking_Diario')
mail.Attachments.Add(str(attachment))

mail.Send()
print('O E-mail da diretoria foi enviado')
    


# In[ ]:




