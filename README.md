# pluggy-sheets-sync

Um script leve em Google Apps Script (GAS) para sincronizar seus dados do Open Finance diretamente para o Google Sheets, utilizando a API da [Pluggy](https://pluggy.ai/).

Crie seu próprio painel financeiro automatizado e *serverless*, rodando 100% na infraestrutura gratuita do Google.

## 🛠 Features

* **Sincronização Completa:** Extrai transações, saldos atuais, limite de cartões de crédito, investimentos e faturas/empréstimos.
* **Idempotente:** Usa os IDs das transações para garantir que nenhuma linha seja duplicada, não importa quantas vezes o script rode.
* **Resiliência:** Implementa *Exponential Backoff* para lidar com `Rate Limits` e instabilidades de rede de forma graciosa.
* **Seguro por padrão:** As chaves da API não ficam *hardcoded*. O script usa o `PropertiesService` do Google para armazenar credenciais sensíveis de forma isolada.
* **Pronto para Análise:** Separa os dados de forma limpa em 4 planilhas (`Transacoes`, `Saldos`, `Investimentos`, `Faturas_e_Emprestimos`), excelente para gerar gráficos ou alimentar modelos de dados.

## 🚀 Quick Start

### 1. Preparando o Ambiente
1. Crie uma nova planilha no [Google Sheets](https://sheets.new/).
2. No menu superior, vá em **Extensões** > **Apps Script**.
3. Apague o código padrão (`myFunction`) e cole todo o conteúdo do arquivo `MeuPluggyIntegration.gs` deste repositório.

### 2. Configurando as Credenciais (Segurança)
Para não deixar seu `CLIENT_ID` e `CLIENT_SECRET` expostos no código:

1. No editor do Apps Script, encontre a função `setupCredentials()` e insira suas chaves provisoriamente:
   ```javascript
   function setupCredentials() {
     PropertiesService.getScriptProperties().setProperties({
       "PLUGGY_CLIENT_ID": "COLE_AQUI",
       "PLUGGY_CLIENT_SECRET": "COLE_AQUI"
     });
   }
   ```
2. *(Opcional)* Se a sua conta Pluggy bloqueia a rota global de listar itens, preencha seus IDs na função `setupManualItemsIds()`. Use um array de UUIDs dos seus itens:
   ```javascript
   function setupManualItemsIds() {
     const ids = [
       "550e8400-e29b-41d4-a716-446655440000",
       "6ba7b810-9dad-41d1-80b4-00c04fd430c8"
     ];
   }
   ```
3. No painel superior, selecione a função **`setupSensitiveGlobalVariables`** e clique em **Executar**.
4. **Apague as chaves do seu código** e salve o arquivo. Suas credenciais agora estão seguras no cofre do projeto.

## 💡 Como Interagir e Usar

O script é projetado para rodar em *background*, mas você pode interagir com ele de duas formas:

**Execução Manual (Para testes ou atualizações imediatas):**
1. Abra o painel do Apps Script.
2. Selecione a função `syncFinances`.
3. Clique em **Executar** (na primeira vez, o Google pedirá que você autorize o script a ler/escrever na sua planilha).
4. As 4 abas serão criadas e populadas automaticamente na sua planilha.

**Execução Automatizada (Cron Job nativo):**
1. No menu esquerdo do Apps Script, clique em **Acionadores** (ícone de relógio).
2. Clique em **Adicionar acionador** (canto inferior direito).
3. Configure da seguinte forma:
   * **Função:** `syncFinances`
   * **Origem do evento:** `Baseado no tempo`
   * **Tipo de acionador:** `Temporizador de horas`
   * **Intervalo:** `A cada 1 hora` (ou 2 horas)
4. Salve. O script fará *polling* da API silenciosamente, injetando novos dados na planilha assim que o Open Finance compensar novas transações.

## 📚 Referências da API Pluggy

* [API Reference](https://docs.pluggy.ai/en/reference)
* [Spec OpenAPI](https://docs.pluggy.ai/openapi/pluggy-api.json): fonte da verdade para os schemas (`Bill`, `Loan`, `Investment`, etc.).
* [Credit Card Bills](https://docs.pluggy.ai/en/docs/products/credit-card-bills): regra de quitação de faturas usada pelo script.

## 📄 Licença
Distribuído sob a licença MIT. Sinta-se livre para fazer um *fork*, modificar e adaptar para as suas próprias regras de negócio ou dashboards!