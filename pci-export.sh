#!/bin/bash

if [ "$EUID" -ne 0 ]; then 
	echo "Should be run as root"; exit 1
fi

FILE_PATH="/opt/pci_dev.text"

lspci -nnmm > "$FILE_PATH"

lspci-docx "$FILE_PATH"
