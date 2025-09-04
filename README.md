# AutoAgendamento_AppScript

Agendamento do usuário via link, com inserção no google agenda da unidade com base em horários e dias definidos em planilha sheets e com check ao google agenda do setor. E mail de confirmação com botão de cancelar disponível ao usuário. Integração com google sheets, agenda google, google meet e gmail.

<img width="1903" height="913" alt="image" src="https://github.com/user-attachments/assets/030493e4-6612-42fa-a19b-ce81e134c5dd" />


# Requisitos Funcionais (RF):

RF01 – O sistema deve permitir que usuários externos realizem agendamentos junto ao setor da UFF.

RF02 – O sistema deve restringir o acesso apenas a pessoas com login institucional id.uff.br.

RF03 – O sistema deve enviar um e-mail de confirmação para cada agendamento realizado.

RF04 – O sistema deve permitir que o usuário cancele um agendamento previamente feito.

RF05 – O sistema deve enviar um e-mail de confirmação sempre que um agendamento for cancelado.

RF06 – O sistema deve registrar automaticamente todos os agendamentos em uma planilha Google Sheets e no Google Agenda.

RF07 – O sistema deve enviar um aviso por e-mail ao usuário agendado com 24h de antecedência do compromisso.

RF08 – O sistema não deve permitir agendamento para alunos que não estão com certificado pronto (base sheets)

RF09 – O sistema deve bloquear agenda em dias com horários indisponíveis e feriados

RF10- O sistema deve identificar assinatura de consentimento de Termo de Privacidade e data de assinatura

# Requisitos Não Funcionais :

RNF01 – O sistema deve estar disponível via navegador web.

RNF02 – O sistema deve possuir interface simples e intuitiva.

RNF03 – O sistema deve seguir padrões de segurança da informação, como sanitizar inputs do lado do servidor, rating limit, confidencialidade com base na conta id.uff.br

RNF04 – O sistema deve seguir a LGPD, garantindo proteção de dados pessoais.
# Fluxo Conceitual
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/15a9e7fe-e7da-48e3-8d8f-48201effadcb" />

# Programas necessários

- VS Code
- GitHub
- Git
- Node
- Clasp
- Planilhas Google
- Google Agenda
- Gmail

## Passo a passo

**BAIXANDO PROJETO DO GITHUB E SUBINDO PARA O APP SCRIPT**
1. **Abrir PowerShell**  
   Clicar em *Windows Powershell* > botão direito > *Executar como administrador*.  
   Verificar se:
   ```
   Get-ExecutionPolicy
   ````
Retorna Restricted.
Se sim, executar:

  ```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
````
2.**Instalar clasp via terminal no VS Code**

   ```
npm install -g @google/clasp
````
3. **Fazer login no clasp**
   
   ```
   clasp login
   ````
   
4. **Clonar projeto do github**
   ```
   git clone https://github.com/RenataVerasVenturim/AutoAgendamento_AppScript.git

   ````

5. **Criar novo projeto App script**
   
   ```
   clasp create --type standalone --title "AutoAgendamento_AppScript"
   ````

6. **Subir código para o projeto App Script criado**

   ```
   clasp push
   ````
7. **Incluir SERVIÇOS AO PROJETO (APIS)**
   <P>
      Com o projeto app script aberto em editor, clicar em "Serviços"> + > Calendar e Sheets
   </P>
8. **Inserir propertiers para proteção de dados de id do calendar google**
<p>Em https://calendar.google.com/, clique nas engrenagens >configurações> Clique na agenda no canto esquerdo embaixo > Em "Integrar agenda"> copie o ID (algo como "abcdefghijk123456789@group.calendar.google.com")
<br> 
<P>No projeto app script, vá em engrenagens> Em "propriedade do script", crie a variável de ambiente "calendarId"=[ID DA AGENDA]
<br> No código, para que o meet seja criado, precisa ser na agenda primary "const calendarId = "primary""</P>

9. **Implantar projeto**
<p>
   Clicar em implantar>Gerenciar Implantações> app web > Versão nova> Inserir descrição> Executar como "Eu (SEU E MAIL)". Quem pode acessar "Qualquer pessoa em Universidade Federal Fluminense"
</p>
   
**ALTERANDO PROJETO QUE ESTÁ NO APP SCRIPT**
  
1. **Clonar projeto do Apps Script** (acesse confiugurações do projeto no app script>código do script)

   ```
   clasp clone [código do projeto no app script]
   ````
   
2.**Enviar alterações para o projeto Apps Script**

   ```
   clasp push
   ````

3.**Enviar alterações para o projeto GitHub** (obs: salvar o readme.me)

   ```
  git init
   ````

   ```
  git remote add origin https://github.com/RenataVerasVenturim/AppScript_Auxilio_Financeiro.git
   ````
```
  git fetch origin
   ````
```
  git add .
 ````
```
  git commit -m "Atualização do projeto Apps Script"
   ````
```
  git branch -M main
 ````
```
  git push -u origin main
   ````


