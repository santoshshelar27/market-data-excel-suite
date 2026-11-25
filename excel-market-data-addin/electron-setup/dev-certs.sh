#!/usr/bin/env bash
set -euo pipefail

CERT_DIR="$(cd "$(dirname "$0")" && pwd)/certs"
mkdir -p "$CERT_DIR"

echo "Generating local Root CA..."
openssl genrsa -out "$CERT_DIR/rootCA.key" 2048
openssl req -x509 -new -nodes -key "$CERT_DIR/rootCA.key" -sha256 -days 3650 \
  -subj "/CN=MarketData Local Root CA" \
  -out "$CERT_DIR/rootCA.crt"

cat > "$CERT_DIR/localhost.ext" <<EOF
authorityKeyIdentifier=keyid,issuer
basicConstraints=CA:FALSE
keyUsage = digitalSignature, nonRepudiation, keyEncipherment, dataEncipherment
subjectAltName = @alt_names

[alt_names]
DNS.1 = localhost
IP.1 = 127.0.0.1
EOF

echo "Generating localhost certificate..."
openssl genrsa -out "$CERT_DIR/localhost.key" 2048
openssl req -new -key "$CERT_DIR/localhost.key" \
  -subj "/CN=localhost" \
  -out "$CERT_DIR/localhost.csr"

openssl x509 -req -in "$CERT_DIR/localhost.csr" \
  -CA "$CERT_DIR/rootCA.crt" -CAkey "$CERT_DIR/rootCA.key" -CAcreateserial \
  -out "$CERT_DIR/localhost.crt" -days 825 -sha256 \
  -extfile "$CERT_DIR/localhost.ext"

rm -f "$CERT_DIR/localhost.csr" "$CERT_DIR/localhost.ext"
echo "Certificates created in $CERT_DIR"
echo "To trust on macOS for local dev, run:"
echo "  sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain \"$CERT_DIR/rootCA.crt\""
