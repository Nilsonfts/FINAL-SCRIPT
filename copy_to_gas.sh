#!/bin/bash
# Скрипт для быстрого копирования файлов в Google Apps Script

echo "=== Список файлов для копирования в Google Apps Script ==="
echo ""

for i in {01..16}; do
    filename=$(ls ${i}_*.gs 2>/dev/null | head -1)
    if [ -f "$filename" ]; then
        echo "Файл: $filename"
        echo "Размер: $(wc -l < "$filename") строк"
        echo "---"
    fi
done

echo ""
echo "Для копирования конкретного файла используйте:"
echo "cat [имя_файла].gs | pbcopy  # на Mac"
echo "cat [имя_файла].gs | xclip -selection clipboard  # на Linux"
echo ""
echo "Затем вставьте в Google Apps Script Editor"
