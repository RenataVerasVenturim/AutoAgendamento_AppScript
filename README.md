# AutoAgendamento_AppScript
Agendamento externo à google agenda da unidade com base em horários e dias definidos e com check google agenda

# Programas necessários

- VS Code
- Node
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
Em https://calendar.google.com/, clique nas engrenagens >configurações> Clique na agenda no canto esquerdo embaixo > Em "Integrar agenda"> copie o ID (algo como "abcdefghijk123456789@group.calendar.google.com")
<br> 
No projeto app script, vá em engrenagens> Em "propriedade do script", crie a variável de ambiente "calendarId"=[ID DA AGENDA]
