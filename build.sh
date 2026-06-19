#!/usr/bin/env bash
#
# Build the StockData LibreOffice Calc add-in (.oxt) on Linux/macOS.
#
# Requires the LibreOffice SDK. Set LO_PROGRAM to the office "program" dir
# (containing types.rdb and unoidl-write) and, for the legacy path, LO_SDK_HOME.
#
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
build="$root/build"
stage="$build/oxt"
dist="$root/dist"

LO_PROGRAM="${LO_PROGRAM:-}"
LO_SDK_HOME="${LO_SDK_HOME:-}"

find_first() { for c in "$@"; do [ -n "$c" ] && [ -e "$c" ] && { echo "$c"; return; }; done; }

if [ -z "$LO_PROGRAM" ]; then
  LO_PROGRAM="$(find_first \
    /usr/lib/libreoffice/program \
    /opt/libreoffice*/program \
    /Applications/LibreOffice.app/Contents/MacOS \
    /Applications/LibreOffice.app/Contents/Frameworks)"
fi
[ -n "$LO_PROGRAM" ] || { echo "LibreOffice program dir not found; set LO_PROGRAM." >&2; exit 1; }

types_rdb="$(find_first "$LO_PROGRAM/types.rdb" "$LO_PROGRAM/types/offapi.rdb")"
[ -n "$types_rdb" ] || { echo "types.rdb not found under $LO_PROGRAM." >&2; exit 1; }

unoidl_write="$(find_first "$LO_PROGRAM/unoidl-write" "${LO_SDK_HOME:-}/bin/unoidl-write")"
idlc="$(find_first "${LO_SDK_HOME:-}/bin/idlc")"

rm -rf "$build" "$dist"
mkdir -p "$stage" "$dist"

idl="$root/src/idl/XStockData.idl"
rdb_out="$stage/StockData.rdb"

if [ -n "$unoidl_write" ]; then
  echo "Compiling IDL with unoidl-write..."
  "$unoidl_write" "$types_rdb" "$idl" "$rdb_out"
elif [ -n "$idlc" ]; then
  echo "Compiling IDL with idlc + regmerge..."
  regmerge="$(find_first "$LO_PROGRAM/regmerge" "${LO_SDK_HOME:-}/bin/regmerge")"
  "$idlc" -w -I "$LO_SDK_HOME/idl" -O "$build" "$idl"
  "$regmerge" "$rdb_out" /UCR "$build/XStockData.urd"
else
  echo "Neither unoidl-write nor idlc found. Install the LibreOffice SDK." >&2
  exit 1
fi

echo "Staging .oxt payload..."
cp "$root/oxt/description.xml" "$stage/"
cp "$root/oxt/CalcAddIns.xcu" "$stage/"
cp "$root/oxt/Jobs.xcu" "$stage/"
cp "$root/src/python/stockdata.py" "$stage/"
cp -r "$root/oxt/META-INF" "$stage/"
cp -r "$root/oxt/description" "$stage/"
cp -r "$root/oxt/icons" "$stage/"
cp -r "$root/oxt/registration" "$stage/"

oxt="$dist/StockData.oxt"
echo "Packing $oxt..."
( cd "$stage" && zip -r -q "$oxt" . )

echo "Done: $oxt"
echo "Install with:  unopkg add --force \"$oxt\""
