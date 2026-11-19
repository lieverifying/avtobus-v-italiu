# Установка сборки для https://avtobus-v-italiu.com из консоли хостинга.
# Запуск:  sh install-adm-tools-com.sh /путь/к/папке/сайта
# Скрипт проверяет архив по SHA-256, делает резервную копию текущего сайта
# и только потом копирует новые файлы. .htaccess и другие файлы хостинга не удаляются.
set -eu
site_dir="${1:?Укажите папку сайта avtobus-v-italiu.com, где лежит index.html}"
backup_dir="${HOME}/site-backups"
test -f "$site_dir/index.html" || { echo "В $site_dir нет index.html — проверьте путь"; exit 1; }
work_dir=$(mktemp -d)
curl -fL --connect-timeout 15 --max-time 120 'https://raw.githubusercontent.com/lieverifying/avtobus-v-italiu/69d5678c33cbde7d7adc5fa48dbe39129b996544/deploy/avtobus-v-italiu-com.zip' -o "$work_dir/site.zip"
printf '%s  %s\n' '97a2aa02744189e6719abec74ff04c438d7a75eb7eba48eddb3ae34c36708517' "$work_dir/site.zip" | sha256sum -c -
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
