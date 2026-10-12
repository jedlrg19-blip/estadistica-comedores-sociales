#!/usr/bin/env bash
# Protege /admin* con Cloudflare Access (Zero Trust) y verifica el despliegue de Pages.
# NO contiene secretos: lea el token de la variable de entorno. SIN PROBAR contra la cuenta real.
#   CLOUDFLARE_API_TOKEN  token con permisos: Account > Access: Apps and Policies > Edit ; Account > Cloudflare Pages > Read
#   ADMIN_EMAILS          correos permitidos, separados por coma (ej. "a@mides.gob.gt,b@mides.gob.gt")
#   PROTECT_DATA=1        (opcional) también protege /data/* (el sitio público dejará de leer los Excel)
set -euo pipefail
ACCOUNT_ID="${ACCOUNT_ID:-220a7358bccaef07bb068c304de87444}"
PROJECT="${PROJECT:-estadistica-comedores-sociales}"
: "${CLOUDFLARE_API_TOKEN:?Defina CLOUDFLARE_API_TOKEN}"; : "${ADMIN_EMAILS:?Defina ADMIN_EMAILS}"
API="https://api.cloudflare.com/client/v4/accounts/$ACCOUNT_ID"
H=(-H "Authorization: Bearer $CLOUDFLARE_API_TOKEN" -H "Content-Type: application/json")
j(){ python3 -c "import sys,json;d=json.load(sys.stdin);print(json.dumps(d,ensure_ascii=False)) if not d.get('success') else print(d['result'] if isinstance(d['result'],str) else d['result'].get('$1',''))"; }

echo "→ Proyecto Pages"; SUB=$(curl -fsS "${H[@]}" "$API/pages/projects/$PROJECT" | python3 -c "import sys,json;print(json.load(sys.stdin)['result']['subdomain'])"); echo "  $SUB"
EMAILS=$(python3 -c "import json,os;print(json.dumps([{'email':{'email':e.strip()}} for e in os.environ['ADMIN_EMAILS'].split(',') if e.strip()]))")
paths=("admin*"); [ "${PROTECT_DATA:-0}" = "1" ] && paths+=("data/*")
for p in "${paths[@]}"; do
  echo "→ Access app para $SUB/$p"
  BODY=$(python3 - <<PY
import json,os
print(json.dumps({"name":"Tablero Comedores · $p","domain":"$SUB/$p","type":"self_hosted","session_duration":"12h",
 "policies":[{"name":"Administradores","decision":"allow","include":$EMAILS}]}))
PY
)
  curl -fsS -X POST "${H[@]}" "$API/access/apps" -d "$BODY" | python3 -c "import sys,json;d=json.load(sys.stdin);print('  ok' if d.get('success') else d)"
done
echo "Listo. Verifique en Zero Trust → Access → Applications."
