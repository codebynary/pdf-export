# 📊 Performance do Extrator de Word

## Tempo de Processamento Esperado

### Por Arquivo
- **Arquivo simples** (1-2 páginas): 1-3 segundos
- **Arquivo médio** (3-5 páginas): 3-7 segundos  
- **Arquivo grande** (5+ páginas): 7-15 segundos

### Fatores que Afetam a Velocidade

1. **Tamanho do arquivo** - Arquivos maiores demoram mais
2. **Número de tabelas** - Mais tabelas = mais tempo
3. **Complexidade da formatação** - Formatação complexa demora mais
4. **Hardware** - CPU e memória RAM disponíveis

## Exemplo de Tempo Total

- **1 arquivo**: ~5 segundos
- **4 arquivos**: ~20-30 segundos
- **10 arquivos**: ~50-90 segundos
- **50 arquivos**: ~4-7 minutos

## O que Fazer se Estiver Muito Lento

### 1. Verificar se está travado
- A barra de progresso está se movendo?
- O log está sendo atualizado?
- Se SIM → Aguarde, está processando
- Se NÃO → Pode estar travado

### 2. Se travou
- Feche a aplicação (botão X)
- Abra novamente
- Tente com menos arquivos primeiro

### 3. Otimizações Possíveis
- Processar em lotes menores (5-10 arquivos por vez)
- Fechar outros programas para liberar memória
- Usar um SSD ao invés de HD (mais rápido)

## Dicas de Uso

✅ **Recomendado:**
- Processar até 20 arquivos por vez
- Aguardar conclusão antes de fechar
- Verificar o log para acompanhar progresso

❌ **Evitar:**
- Processar 100+ arquivos de uma vez
- Fechar a aplicação durante processamento
- Abrir os arquivos Word enquanto processa
