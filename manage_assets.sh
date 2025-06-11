# manage_assets.sh - Unix wrapper for CityInfraXLS

python3 manage_assets.py "$@"
exit_code=$?

if [ $exit_code -ne 0 ]; then
  echo ""
  echo "Command failed with error code $exit_code"
fi