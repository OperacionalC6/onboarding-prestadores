#!/usr/bin/env bash
# Prepara o GCP para o Onboarding de Prestadores guardar dados (Firestore) e anexos (Cloud Storage).
# Rode no Cloud Shell (console.cloud.google.com, ícone ">_" no topo):
#   curl -sL https://raw.githubusercontent.com/OperacionalC6/onboarding-prestadores/main/setup_gcp.sh | bash -s ID_DO_PROJETO
# No fim ele mostra o bloco para colar em "Manage app > Settings > Secrets" no Streamlit.
set -uo pipefail

# Tudo dentro de main(): com "curl | bash", o bash só executa depois de ler o script inteiro;
# senão o gcloud consome o resto do script pela entrada padrão.
main() {
  echo "== Projetos disponíveis =="
  gcloud projects list --format="table(projectId,name)"
  PROJ="${1:-}"
  if [ -z "$PROJ" ]; then
    read -r -p "Digite o ID do projeto (coluna PROJECT_ID): " PROJ < /dev/tty
  fi
  gcloud config set project "$PROJ" >/dev/null || exit 1

  SA_NOME=onboarding-app
  SA="$SA_NOME@$PROJ.iam.gserviceaccount.com"
  BUCKET="onboarding-anexos-$PROJ"

  echo "== Ativando as APIs (pode levar 1 minuto) =="
  # Projeto novo às vezes responde "Quota exceeded" (limite por minuto): espera e tenta de novo.
  for tentativa in 1 2 3 4 5 6; do
    gcloud services enable firestore.googleapis.com storage.googleapis.com iam.googleapis.com 2>/tmp/enable.err && break
    if [ "$tentativa" = 6 ] || ! grep -q "429\|RATE_LIMIT\|Quota exceeded" /tmp/enable.err; then
      cat /tmp/enable.err; exit 1
    fi
    echo "Google pediu para aguardar (limite por minuto). Tentando de novo em 60s... ($tentativa/5)"
    sleep 60
  done

  echo "== Banco Firestore (São Paulo) =="
  if gcloud firestore databases describe --database="(default)" >/dev/null 2>&1; then
    echo "Já existe, mantendo."
  else
    gcloud firestore databases create --location=southamerica-east1 --type=firestore-native || exit 1
  fi

  echo "== Bucket de anexos =="
  if gcloud storage buckets describe "gs://$BUCKET" >/dev/null 2>&1; then
    echo "Já existe, mantendo."
  elif ! gcloud storage buckets create "gs://$BUCKET" --location=us-east1 --uniform-bucket-level-access \
         --public-access-prevention; then
    echo ""
    echo ">> Não consegui criar o bucket. Normalmente é porque o projeto está sem faturamento (cartão)."
    echo ">> Ative em https://console.cloud.google.com/billing/linkedaccount?project=$PROJ e rode de novo."
    exit 1
  fi

  echo "== Conta de serviço do app =="
  gcloud iam service-accounts describe "$SA" >/dev/null 2>&1 || \
    gcloud iam service-accounts create "$SA_NOME" --display-name="Onboarding Streamlit" || exit 1
  sleep 5
  gcloud projects add-iam-policy-binding "$PROJ" --member="serviceAccount:$SA" --role=roles/datastore.user \
    --condition=None -q >/dev/null || exit 1
  gcloud storage buckets add-iam-policy-binding "gs://$BUCKET" --member="serviceAccount:$SA" \
    --role=roles/storage.objectAdmin >/dev/null || exit 1

  CHAVE=$(mktemp)
  if ! gcloud iam service-accounts keys create "$CHAVE" --iam-account="$SA"; then
    echo ">> A organização bloqueia a criação de chaves (política iam.disableServiceAccountKeyCreation)."
    echo ">> Peça ao administrador do GCP uma exceção para este projeto e rode de novo."
    rm -f "$CHAVE"; exit 1
  fi

  echo ""
  echo "================= COPIE DAQUI PARA BAIXO ================="
  python3 -c '
import json, sys
chave, bucket = json.load(open(sys.argv[1])), sys.argv[2]
print("[armazenamento]")
print("firestore = true")
print(f"gcs_bucket = \"{bucket}\"")
print()
print("[gcp_service_account]")
for k, v in chave.items():
    print(f"{k} = {json.dumps(v)}")
' "$CHAVE" "$BUCKET"
  echo "================= ATÉ AQUI ================="
  rm -f "$CHAVE"
  echo "Cole no fim do Secrets do Streamlit (Manage app > Settings > Secrets) e salve."
}

main "$@" < /dev/null
