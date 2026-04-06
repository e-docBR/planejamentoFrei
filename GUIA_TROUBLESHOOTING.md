# 🔧 Guia de Troubleshooting - Sistema de Planejamento v9.20

## 📋 Índice
- [Ferramentas de Diagnóstico](#ferramentas-de-diagnóstico)
- [Problemas Comuns e Soluções](#problemas-comuns-e-soluções)
- [Como Executar o Diagnóstico](#como-executar-o-diagnóstico)
- [Interpretando os Resultados](#interpretando-os-resultados)
- [Checklist de Instalação](#checklist-de-instalação)

---

## 🛠️ Ferramentas de Diagnóstico

O sistema v9.20 inclui uma função de diagnóstico automático que testa:

✅ Autenticação do usuário  
✅ Configuração do sistema (CONFIG)  
✅ Acesso à pasta do Google Drive  
✅ Criação/acesso da planilha de registros  
✅ Cadastros de professores  
✅ Perfil e permissões do usuário  

---

## 🚨 Problemas Comuns e Soluções

### 1️⃣ "Planilha não foi criada"

**Sintomas:**
- Sistema executa mas planilha não aparece no Drive
- Erro ao tentar acessar dados

**Causas possíveis:**
- ID da pasta incorreto
- Falta de permissão no Drive
- Erro durante a criação

**Soluções:**
1. Execute `diagnosticoSistema()` (veja abaixo como executar)
2. Verifique o teste #3 (ACESSO À PASTA)
3. Verifique o teste #4 (PLANILHA)
4. Se o ID da pasta estiver errado:
   - Abra a pasta no Drive
   - Copie o ID da URL (parte após `/folders/`)
   - Cole em `CONFIG.PASTA_ID`

---

### 2️⃣ "Menus não funcionam"

**Sintomas:**
- Sistema carrega mas nenhum menu responde
- Página fica em branco ou sem interação

**Causas possíveis:**
- Erro JavaScript no navegador
- Sistema não foi implantado como Web App
- Perfil não foi reconhecido

**Soluções:**
1. **Verifique o console do navegador:**
   - Pressione `F12` para abrir DevTools
   - Vá na aba "Console"
   - Veja se há erros em vermelho
   - Copie os erros e analise

2. **Verifique se foi feito deploy correto:**
   - Apps Script Editor → "Implantar" → "Implantar como Web App"
   - Certifique-se de que está usando a URL do Web App, não a URL do editor

3. **Execute o diagnóstico:**
   - Teste #6 mostra se seu perfil está autorizado
   - Se aparecer "desconhecido", seu email não está cadastrado

---

### 3️⃣ "Acesso negado"

**Sintomas:**
- Página mostra "Acesso não autorizado"
- Email não reconhecido

**Causas:**
- Email não está em `CONFIG.EMAILS_COORDENACAO`
- Email não está cadastrado na aba Cadastros

**Soluções:**
1. **Para coordenação:**
   - Adicione o email em `CONFIG.EMAILS_COORDENACAO`
   - Exemplo: `EMAILS_COORDENACAO: ["seu.email@edu.mucuri.ba.gov.br"]`

2. **Para professores:**
   - A coordenação deve cadastrar o professor na aba "Cadastros"
   - Use exatamente o mesmo email que o professor usa para fazer login

---

### 4️⃣ "Erro ao criar planilha"

**Sintomas:**
- Mensagem de erro durante primeiro acesso
- Sistema não consegue inicializar

**Causas:**
- Falta de permissão para criar planilhas
- Quota de armazenamento excedida
- Problema de conexão

**Soluções:**
1. **Verifique permissões:**
   - Na primeira execução, o Google solicita permissões
   - Clique em "Avançado" → "Ir para [projeto]" → "Permitir"

2. **Verifique espaço no Drive:**
   - Vá em drive.google.com
   - Veja se tem espaço disponível
   - Libere espaço se necessário

3. **Execute o diagnóstico:**
   - Teste #4 mostrará o erro exato

---

### 5️⃣ "Sistema muito lento"

**Causas:**
- Muitos registros na planilha
- Cache desativado
- Conexão lenta

**Soluções:**
1. O sistema usa cache automático de 5 minutos
2. Evite abrir muitas abas simultaneamente
3. Considere arquivar registros antigos

---

## 📊 Como Executar o Diagnóstico

### Passo a Passo:

1. **Abra o Apps Script Editor:**
   - Vá para [script.google.com](https://script.google.com)
   - Abra o projeto "Sistema de Planejamento Escolar"

2. **Selecione a função診断:**
   - No topo da tela, há um dropdown com lista de funções
   - Selecione `diagnosticoSistema`

3. **Execute:**
   - Clique no ícone de "play" (▶️) ao lado do dropdown
   - Aguarde alguns segundos

4. **Veja os resultados:**
   - Clique em "Execuções" no menu lateral esquerdo
   - Clique na execução mais recente
   - Role o log para ver o relatório completo

---

## 🔍 Interpretando os Resultados

O diagnóstico mostra 6 testes:

### ✅ Símbolo Verde = Teste passou
Tudo certo nesta área!

### ⚠️  Símbolo Amarelo = Aviso
Não é erro crítico, mas requer atenção.

**Exemplos:**
- "Nenhum professor cadastrado ainda" → Normal no primeiro acesso
- "Abas faltando" → Serão criadas automaticamente

### ❌ Símbolo Vermelho = Erro crítico
Precisa ser corrigido imediatamente.

**Exemplos:**
- "Email inválido ou não detectado"
- "PASTA_ID não configurado"
- "Erro ao acessar pasta"

---

## ✅ Checklist de Instalação

Use este checklist para garantir que tudo está configurado:

### Configuração Inicial:
- [ ] `CONFIG.PASTA_ID` preenchido com ID da pasta do Drive
- [ ] `CONFIG.EMAILS_COORDENACAO` contém pelo menos um email
- [ ] `CONFIG.ANO_LETIVO` configurado (ex: "2025")
- [ ] `CONFIG.FUSO` configurado (ex: "America/Bahia")

### Deploy:
- [ ] Feito deploy como "Web App"
- [ ] Configurado "Executar como: Eu"
- [ ] Configurado "Quem tem acesso: Qualquer pessoa"
- [ ] URL do Web App copiada e acessível

### Permissões:
- [ ] Permissões do Google autorizadas
- [ ] Acesso ao Drive concedido
- [ ] Acesso ao Sheets concedido
- [ ] Acesso ao Gmail concedido (para envio de emails)

### Pasta do Drive:
- [ ] Pasta existe e está acessível
- [ ] Você tem permissão de Editor na pasta
- [ ] ID da pasta está correto em `CONFIG.PASTA_ID`

### Primeiro Acesso:
- [ ] Acessando via URL do Web App (não URL do editor)
- [ ] Email usado é o mesmo configurado em `EMAILS_COORDENACAO`
- [ ] Sistema executou `diagnosticoSistema()` sem erros críticos

---

## 🆘 Suporte Adicional

### Se o problema persistir:

1. **Execute o diagnóstico completo:**
   ```
   diagnosticoSistema()
   ```

2. **Capture o relatório completo:**
   - Copie TODO o texto do log
   - Inclua Stack Traces se houver erros

3. **Verifique os logs do sistema:**
   - Apps Script Editor → "Execuções"
   - Veja se há erros recentes
   - Anote timestamp e mensagens de erro

4. **Informações úteis para debug:**
   - Versão do navegador
   - Email que está tentando acessar
   - URL exata que está sendo usada
   - Hora exata do erro
   - Mensagem de erro completa

---

## 📝 Funções Auxiliares

### `infosPlanilha()`
Mostra informações sobre a planilha atual:
- Nome
- ID
- URL
- Número de abas
- Quantidade de registros

### `resetarIdPlanilha()`
Remove o ID salvo da planilha. Use se:
- Quer forçar o sistema a buscar/criar nova planilha
- ID salvo está incorreto
- Quer "resetar" o sistema

⚠️ **Atenção:** Isso não exclui a planilha, apenas faz o sistema esquecê-la.

---

## 🎯 Fluxo de Resolução de Problemas

```
Problema relatado
    ↓
Execute diagnosticoSistema()
    ↓
Veja qual teste falhou (❌)
    ↓
┌─────────────────┬──────────────────┬─────────────────┐
│ Teste #1 falhou │ Teste #3 falhou  │ Teste #4 falhou │
│ (Autenticação)  │ (Pasta Drive)    │ (Planilha)      │
├─────────────────┼──────────────────┼─────────────────┤
│ Verifique email │ Verifique PASTA_ID│ Veja erro exato │
│ Configure em    │ Verifique permis. │ Verifique permis│
│ EMAILS_COORDENACAO│ Verifique existência│ Execute novamente│
└─────────────────┴──────────────────┴─────────────────┘
    ↓
Corrija o problema identificado
    ↓
Execute diagnosticoSistema() novamente
    ↓
Todos os testes ✅? → Sistema funcionando!
```

---

## 🔄 Versão
- **Versão do Sistema:** v9.20
- **Data:** 2025
- **Última atualização deste guia:** Data de hoje

---

**💡 Dica:** Mantenha este guia salvo junto com o código do projeto para referência rápida!
