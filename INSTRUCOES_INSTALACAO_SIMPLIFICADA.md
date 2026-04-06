# 📘 INSTRUÇÕES DE INSTALAÇÃO SIMPLIFICADA
## Sistema de Planejamento Escolar v9 - Criação Automática de Planilha

---

## ✨ NOVIDADE v9.20: Instalação Ainda Mais Fácil!

A partir da versão 9.20, **você não precisa mais criar a planilha manualmente** e nem copiar o ID!

A planilha de registros será **criada automaticamente** no primeiro acesso ao sistema.

---

## 🚀 CONFIGURAÇÃO MÍNIMA NECESSÁRIA

Agora você só precisa configurar **2 itens** no arquivo `Codigo_v9.gs`:

### 1️⃣ **PASTA_ID** (obrigatório)
```javascript
PASTA_ID: "COLE_AQUI_O_ID_DA_PASTA_DO_DRIVE"
```

**Como obter:**
1. Crie uma pasta no Google Drive para o sistema
2. Abra a pasta
3. Copie o ID da URL: `https://drive.google.com/drive/u/0/folders/ID_ESTÁ_AQUI`
4. Cole no código

### 2️⃣ **EMAILS_COORDENACAO** (obrigatório)
```javascript
EMAILS_COORDENACAO: [
  "coordenadora@escola.edu.br",
  // Adicione mais emails se necessário
]
```

### 3️⃣ **PLANILHA_ID** (opcional - pode deixar como está!)
```javascript
PLANILHA_ID: "COLE_AQUI_O_ID_DA_PLANILHA_REGISTRO",  // OPCIONAL
```

✅ **Pode deixar com placeholder!** A planilha será criada automaticamente.

---

## 📋 PASSO A PASSO COMPLETO

### 1. Configure as 2 informações obrigatórias
- ID da pasta do Drive
- Email(s) da coordenação

### 2. Publique o Web App
1. No editor do Apps Script → **Implantar** → **Nova implantação**
2. Tipo: **Aplicativo da Web**
3. Executar como: **Eu**
4. Quem tem acesso: **Qualquer pessoa**
5. Clique em **Implantar**
6. Copie a URL gerada

### 3. Acesse o sistema pela primeira vez
- Abra a URL do Web App
- **A planilha será criada automaticamente!**
- O ID será salvo automaticamente nas propriedades do script

### 4. Pronto! ✨
O sistema está funcionando!

---

## 🔧 FUNÇÕES ÚTEIS NO EDITOR

Execute estas funções manualmente no editor do Apps Script quando necessário:

### `infosPlanilha()`
Exibe informações sobre a planilha de registros
```
Mostra:
- ID configurado no código
- ID salvo automaticamente
- ID que está sendo usado
- Nome e URL da planilha
- Abas existentes
```

### `resetarIdPlanilha()`
Remove o ID salvo automaticamente
```
Use quando:
- Quiser criar uma nova planilha do zero
- O ID salvo estiver inválido
- Precisar forçar uma nova criação
```

---

## 🎯 COMO FUNCIONA

### Ordem de Busca do ID da Planilha:
1. **CONFIG.PLANILHA_ID** (se configurado e válido)
2. **Propriedades do Script** (ID salvo automaticamente)
3. **Busca por nome** na pasta configurada
4. **Cria nova planilha** se não encontrar

### O que o sistema faz automaticamente:
✅ Busca planilha existente pelo nome  
✅ Cria nova planilha se não encontrar  
✅ Move a planilha para a pasta correta  
✅ Salva o ID nas propriedades do script  
✅ Loga informações úteis no console  

---

## 📊 EXEMPLO DE LOG NA PRIMEIRA EXECUÇÃO

```
📝 Criando planilha de registros: Registro_Planejamentos_2026
✓ ID da planilha salvo automaticamente: 1AbCdEfGhIjKlMnOpQrStUvWxYz
💡 OPCIONAL: Para melhor performance, defina CONFIG.PLANILHA_ID = "1AbC..."

✅ PLANILHA CRIADA COM SUCESSO!
📋 Nome: Registro_Planejamentos_2026
🆔 ID: 1AbCdEfGhIjKlMnOpQrStUvWxYz
📁 Pasta: Planejamentos 2026

ℹ️  O ID foi salvo automaticamente. O sistema está pronto para uso!
```

---

## ⚡ MELHOR PERFORMANCE (OPCIONAL)

Embora não seja obrigatório, você pode **copiar o ID gerado** e colar no código para melhor performance:

1. Execute `infosPlanilha()` no editor
2. Copie o ID exibido
3. Cole em `CONFIG.PLANILHA_ID`

**Vantagem:** Evita busca por nome a cada execução (milissegundos mais rápido)

---

## 🆘 RESOLUÇÃO DE PROBLEMAS

### "Não foi possível acessar a planilha de registros"
- ✅ Verifique se PASTA_ID está correto
- ✅ Verifique permissões da pasta no Drive
- ✅ Execute `resetarIdPlanilha()` e tente novamente

### "Planilha criada mas não aparece na pasta"
- ✅ Aguarde alguns segundos e atualize o Drive
- ✅ Verifique se o ID da pasta está correto
- ✅ Verifique permissões da pasta

### "Quer usar planilha existente de outro ano"
1. Renomeie a planilha antiga no Drive
2. Execute `resetarIdPlanilha()`
3. Acesse o sistema (criará nova planilha com nome do ano atual)

### "Quer forçar criação de nova planilha"
1. Execute `resetarIdPlanilha()` no editor
2. Exclua ou renomeie a planilha antiga no Drive
3. Acesse o sistema novamente

---

## 📝 COMPARAÇÃO COM VERSÃO ANTERIOR

| Aspecto | Versão Anterior | Nova Versão (v9.20) |
|---------|----------------|---------------------|
| Criar planilha manualmente | ✅ Obrigatório | ❌ Opcional |
| Copiar ID da planilha | ✅ Obrigatório | ❌ Opcional |
| Colar ID no código | ✅ Obrigatório | ❌ Opcional |
| Busca automática | ❌ Não tinha | ✅ Inteligente |
| Criação automática | ❌ Não tinha | ✅ Completa |
| Salvamento do ID | ❌ Manual | ✅ Automático |

---

## ✅ CHECKLIST DE INSTALAÇÃO

- [ ] Criar pasta no Google Drive
- [ ] Copiar ID da pasta
- [ ] Colar em `CONFIG.PASTA_ID`
- [ ] Configurar `EMAILS_COORDENACAO`
- [ ] Publicar Web App
- [ ] Acessar URL do Web App
- [ ] ✨ Sistema funcionando!

**Tempo estimado:** 5 minutos

---

## 💡 DICAS PRO

1. **Backup automático:** A planilha fica na pasta do Drive, facilitando backups
2. **Múltiplos anos:** Altere `CONFIG.ANO_LETIVO` para criar planilha de novo ano
3. **Nome único:** Cada ano tem sua própria planilha (Registro_Planejamentos_2026)
4. **Sem conflitos:** Sistema busca antes de criar, evitando duplicatas

---

## 🎉 CONCLUSÃO

Com a versão 9.20, a instalação ficou **extremamente simples**:
- Configure PASTA_ID
- Configure emails
- Publique
- Acesse

**E pronto!** A planilha é criada automaticamente. 🚀

---

*Sistema de Planejamento Escolar v9.20*  
*Colégio Municipal de 1º e 2º Graus de Itabatan*  
*Última atualização: Abril/2026*
