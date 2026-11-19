# Установка сборки для https://avtobus-v-italiu.com из консоли хостинга.
# Запуск:  sh install-adm-tools-com.sh            (папка найдётся сама)
#     или: sh install-adm-tools-com.sh /путь/к/папке/сайта
# Скрипт проверяет архив по SHA-256, делает резервную копию текущего сайта
# и только потом копирует новые файлы. .htaccess и другие файлы хостинга не удаляются.
set -eu
site_dir="${1:-}"
if [ -z "$site_dir" ]; then
  # без аргумента: шукаємо єдину папку, де index.html має canonical avtobus-v-italiu.com
  found=$(grep -rl --include=index.html 'rel="canonical" href="https://avtobus-v-italiu.com/"' "$HOME" 2>/dev/null | grep -v -e '/ru/' -e '/it/' -e site-backups || true)
  n=$(printf '%s\n' "$found" | grep -c . || true)
  if [ "$n" != 1 ]; then echo "Знайдено папок: $n — вкажіть папку сайту аргументом:"; printf '%s\n' "$found"; exit 1; fi
  site_dir=$(dirname "$found")
fi
echo "Папка сайту: $site_dir"
backup_dir="${HOME}/site-backups"
test -f "$site_dir/index.html" || { echo "В $site_dir нет index.html — проверьте путь"; exit 1; }
work_dir=$(mktemp -d)
curl -fL --connect-timeout 15 --max-time 120 'https://raw.githubusercontent.com/lieverifying/avtobus-v-italiu/5e51a7376b56d9cd5890c0f1f5a7e495d3099633/deploy/avtobus-v-italiu-com.zip' -o "$work_dir/site.zip"
printf '%s  %s\n' '044e0066753f63ef58a3ecda4c129f16f8fcda1b02a9a1ccd1e31facac3756f5' "$work_dir/site.zip" | sha256sum -c -
unzip -q "$work_dir/site.zip" -d "$work_dir/new"
grep -q 'https://avtobus-v-italiu.com/' "$work_dir/new/index.html"
mkdir -p "$backup_dir"
backup_file="$backup_dir/avtobus-com-$(date +%Y%m%d-%H%M%S).tar.gz"
tar -czf "$backup_file" -C "$site_dir" .
tar -tzf "$backup_file" > /dev/null
find "$work_dir/new" -type d -exec chmod 755 {} +
find "$work_dir/new" -type f -exec chmod 644 {} +
cp -Rp "$work_dir/new/." "$site_dir/"
rm -rf "$work_dir"
printf '\nГОТОВО. Резервная копия: %s\nОткат: tar -xzf %s -C %s\n' "$backup_file" "$backup_file" "$site_dir"
