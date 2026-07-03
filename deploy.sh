#!/usr/bin/env bash
# ExcportCuy deploy helper
# Usage:
#   ./deploy.sh up                  # build & start (app only) on port 9812
#   ./deploy.sh up --with-db        # sekaligus start MySQL lokal (profile db)
#   ./deploy.sh down                # stop & remove containers
#   ./deploy.sh down --with-db      # stop app + db
#   ./deploy.sh restart [--with-db]
#   ./deploy.sh rebuild [--with-db] # force rebuild image lalu up
#   ./deploy.sh logs [service]      # tail logs (default: excportcuy)
#   ./deploy.sh status
#   ./deploy.sh clean [--with-db]   # down + hapus volume + image
#   ./deploy.sh db:shell            # masuk ke mysql shell (butuh db aktif)
#
# Data lokal:
#   ./data/exports  -> hasil DOCX
#   ./data/uploads  -> file HTML yang di-upload
#   ./data/mysql    -> data MySQL (kalau --with-db)

set -euo pipefail
cd "$(dirname "$0")"

# --- pilih docker compose command ---
if docker compose version >/dev/null 2>&1; then
  DC="docker compose"
elif command -v docker-compose >/dev/null 2>&1; then
  DC="docker-compose"
else
  echo "❌ docker compose tidak ditemukan. Install Docker dulu."
  exit 1
fi

# --- parse args: cari flag --with-db di posisi manapun ---
WITH_DB=0
POSITIONAL=()
for arg in "$@"; do
  case "$arg" in
    --with-db|-d) WITH_DB=1 ;;
    *) POSITIONAL+=("$arg") ;;
  esac
done
set -- "${POSITIONAL[@]:-}"
CMD="${1:-up}"
SUB="${2:-}"

# --- siapkan folder data lokal ---
mkdir -p ./data/exports ./data/uploads
if [[ "$WITH_DB" -eq 1 ]]; then
  mkdir -p ./data/mysql
fi

# --- compose invocation dengan / tanpa profile db ---
COMPOSE=($DC)
if [[ "$WITH_DB" -eq 1 ]]; then
  COMPOSE+=(--profile db)
  export DATABASE_URL="mysql://excport:excport@db:3306/excportcuy"
else
  export DATABASE_URL=""
fi

banner() { echo -e "🐉  $*"; }

case "$CMD" in
  up)
    if [[ "$WITH_DB" -eq 1 ]]; then
      banner "Menyalakan ExcportCuy + MySQL di http://localhost:9812 ..."
    else
      banner "Menyalakan ExcportCuy (tanpa DB) di http://localhost:9812 ..."
    fi
    "${COMPOSE[@]}" up -d --build
    echo "✅  Running. Cek: ./deploy.sh status"
    if [[ "$WITH_DB" -eq 1 ]]; then
      echo "🗄   MySQL: localhost:3306  user=excport pass=excport  db=excportcuy"
    fi
    ;;
  down)
    banner "Menghentikan ExcportCuy ..."
    "${COMPOSE[@]}" down
    ;;
  restart)
    "${COMPOSE[@]}" restart
    ;;
  rebuild)
    banner "Force rebuild ..."
    "${COMPOSE[@]}" build --no-cache
    "${COMPOSE[@]}" up -d
    ;;
  logs)
    svc="${SUB:-excportcuy}"
    "${COMPOSE[@]}" logs -f --tail=200 "$svc"
    ;;
  status|ps)
    "${COMPOSE[@]}" ps
    ;;
  clean)
    banner "Menghapus container, volume, dan image ..."
    "${COMPOSE[@]}" down -v --rmi local || true
    ;;
  db:shell)
    banner "Masuk MySQL shell (excportcuy) ..."
    $DC exec db mysql -uexcport -pexcport excportcuy
    ;;
  *)
    echo "Usage: $0 {up|down|restart|rebuild|logs|status|clean|db:shell} [--with-db]"
    exit 1
    ;;
esac
