#!/bin/bash
# Автоматическая загрузка всех модулей в Google Apps Script

echo "=== АВТОМАТИЧЕСКАЯ ЗАГРУЗКА МОДУЛЕЙ ==="
echo "Проект ID: 1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH"
echo ""

# Проверяем аутентификацию
echo "Проверяем подключение к Google Apps Script..."
clasp list

if [ $? -eq 0 ]; then
    echo "✅ Аутентификация успешна!"
    
    echo "Загружаем все файлы..."
    clasp push --force
    
    echo "✅ Все файлы загружены в Google Apps Script!"
    echo ""
    echo "Откройте проект: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit"
    
else
    echo "❌ Ошибка аутентификации. Необходимо выполнить 'clasp login' сначала."
    echo "Используйте альтернативный метод копирования файлов:"
    echo ""
    
    for i in {01..16}; do
        filename=$(ls ${i}_*.gs 2>/dev/null | head -1)
        if [ -f "$filename" ]; then
            echo "=== $filename ==="
            echo "Размер: $(wc -l < "$filename") строк"
            echo "Команда для копирования: cat '$filename'"
            echo ""
        fi
    done
fi
