set -eu
site_dir='/home/xk589064/dg-s.space/avtobus-v-italiu'
backup_dir='/home/xk589064/site-backups'
test -f "$site_dir/index.html"
work_dir=$(mktemp -d)
curl -fL --connect-timeout 15 --max-time 120 'https://raw.githubusercontent.com/lieverifying/avtobus-v-italiu/3baf3f91adfa38565ddc423e793aa757fbac27fa/deploy/avtobus-v-italiu-ready.zip' -o "$work_dir/site.zip"
printf '%s  %s\n' 'fe71e61f9a2fd540b413b8ca5f9185be0da2cc130ead9b3b264917097b95cd5e' "$work_dir/site.zip" | sha256sum -c -
unzip -q "$work_dir/site.zip" -d "$work_dir/new" -x DEPLOY.txt
mkdir -p "$backup_dir"
backup_file="$backup_dir/avtobus-$(date +%Y%m%d-%H%M%S).tar.gz"
tar -czf "$backup_file" -C "$site_dir" .
tar -tzf "$backup_file" > /dev/null
find "$work_dir/new" -type d -exec chmod 755 {} +
find "$work_dir/new" -type f -exec chmod 644 {} +
cp -Rp "$work_dir/new/." "$site_dir/"
printf '\nГОТОВО. Резервная копия: %s\n' "$backup_file"
