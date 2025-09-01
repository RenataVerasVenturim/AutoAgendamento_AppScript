# AutoAgendamento_AppScript

Agendamento do usuário via link, com inserção no google agenda da unidade com base em horários e dias definidos em planilha sheets e com check ao google agenda do setor. E mail de confirmação com botão de cancelar disponível ao usuário. Integração com google sheets, agenda google, google meet e gmail.

<img width="1365" height="630" alt="image" src="https://github.com/user-attachments/assets/00896768-0758-4265-9230-c22fab51268a" />

# Requisitos
- Usuários externos com total autonomia para realizar agendamentos junto ao setor da UFF
- Apenas pessoas com id.uff.br podem acessar a aplicação
- Os agendamentos devem gerar uma confirmação por e mail
- Os agendamento podem ser cancelados pelo usuário
- Os agendamentos cancelados devem gerar uma confirmação de cancelamento por e mail
- Os agendamentos são registrados automaticamente na planilhas Google e google agenda
- O usuário agendados deve receber aviso 24h antes da data marcada

# Programas necessários

- VS Code
- GitHub
- Git
- Clasp

## Passo a passo

**BAIXANDO PROJETO DO GITHUB E SUBINDO PARA O APP SCRIPT**
1. **Abrir PowerShell**  
   Clicar em *Windows* > botão direito > *Executar como administrador*.  
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

**ALTERANDO PROJETO QUE ESTÁ NO APP SCRIPT**
  
4. **Clonar projeto do Apps Script** (acesse confiugurações do projeto no app script>código do script)

   ```
   clasp clone [código do projeto no app script]
   ````
   
5.**Enviar alterações para o projeto Apps Script**

   ```
   clasp push
   ````

6.**Enviar alterações para o projeto GitHub** (obs: salvar o readme.me)

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

**Inserir propertiers para proteção de dados de id do calendar google**
<p>Em https://calendar.google.com/, clique nas engrenagens >configurações> Clique na agenda no canto esquerdo embaixo > Em "Integrar agenda"> copie o ID (algo como "abcdefghijk123456789@group.calendar.google.com")
<br> 
No projeto app script, vá em engrenagens> Em "propriedade do script", crie a variável de ambiente "calendarId"=[ID DA AGENDA]
<br> No código, para que o meet seja criado, precisa ser na agenda primary "const calendarId = "primary""
