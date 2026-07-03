#!/usr/bin/env bash
# ExcportCuy deploy helper
# Usage:
#   ./deploy.sh up        # build & start (detached) on port 9812
#   ./deploy.sh down      # stop & remove containers
#   ./deploy.sh restart   # restart service
#   ./deploy.sh rebuild   # force rebuild image then up
#   ./deploy.sh logs      # tail logs
#   ./deploy.sh status    # ps
#   ./deploy.sh clean     # down + remove volumes + prune image
#
# Data lokal: aplikasi ini murni client-side (parse HTML -> Word di browser),
# tidak butuh database. Kalau nanti perlu MySQL lokal, uncomment service `db`
# di docker-compose.yml — data akan tersimpan di ./data/mysql (bind mount lokal).

set -euo pipefail

cd "$(dirname "$0")"

# pilih perintah compose yang tersedia
if docker compose version >/dev/null 2>&1; then
  DC="docker compose"
elif command -v docker-compose >/dev/null 2>&1; then
  DC="docker-compose"
else
  echo "❌ docker compose tidak ditemukan. Install Docker dulu."
  exit 1
fi

# pastikan folder data lokal ada (dipakai kalau MySQL diaktifkan)
mkdir -p ./data/mysql ./data/uploads

CMD="${1:-up}"

case "$CMD" in
  up)
    echo "🐉  Building & starting ExcportCuy on http://localhost:9812 ..."
    $DC up -d --build
    echo "✅  Running. Cek: $DC ps"
    ;;
  down)
    echo "🛑  Stopping ExcportCuy ..."
    $DC down
    ;;
  restart)
    $DC restart
    ;;
  rebuild)
    echo "♻️   Force rebuild ..."
    $DC build --no-cache
    $DC up -d
    ;;
  logs)
    $DC logs -f --tail=200
    ;;
  status|ps)
    $DC ps
    ;;
  clean)
    echo "🧹  Removing containers, volumes, and image ..."
    $DC down -v --rmi local || true
    ;;
  *)
    echo "Usage: $0 {up|down|restart|rebuild|logs|status|clean}"
    exit 1
    ;;
esac
