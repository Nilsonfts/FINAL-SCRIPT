#!/bin/bash
# Создание готовых файлов для копирования в Google Apps Script

echo "🚀 ПОДГОТОВКА МОДУЛЕЙ ДЛЯ GOOGLE APPS SCRIPT"
echo "=============================================="
echo ""

# Создаем директорию для готовых файлов
mkdir -p ready_modules

# Список модулей с описаниями
declare -A modules=(
    ["01_Главный_Скрипт.gs"]="🎯 Главный Скрипт - Центральная точка управления"
    ["02_Конфигурация.gs"]="⚙️ Конфигурация - Настройки системы и API ключи"
    ["03_Основные_Утилиты.gs"]="🔧 Основные Утилиты - Базовые функции"
    ["04_Утилиты_Дат.gs"]="📅 Утилиты Дат - Работа с датами"
    ["05_Утилиты_Заголовков.gs"]="📝 Утилиты Заголовков - Обработка метаданных"
    ["06_Утилиты_UTM.gs"]="🔗 Утилиты UTM - Анализ UTM меток"
    ["07_Загрузка_Данных.gs"]="📥 Загрузка Данных - Импорт из источников"
    ["08_Анализ_Каналов.gs"]="📊 Анализ Каналов - Эффективность трафика"
    ["09_Анализ_Лидов.gs"]="👥 Анализ Лидов - Обработка лидов"
    ["10_UTM_Аналитика.gs"]="📈 UTM Аналитика - Детальная UTM аналитика"
    ["11_Первые_Касания.gs"]="🎯 Первые Касания - Анализ касаний"
    ["12_Ежедневная_Статистика.gs"]="📊 Ежедневная Статистика - Дневные отчеты"
    ["13_Эффективность_Менеджеров.gs"]="👨‍💼 Эффективность Менеджеров - Анализ работы"
    ["14_Сравнение_Месяцев.gs"]="📅 Сравнение Месяцев - Сравнительный анализ"
    ["15_Сводка_AmoCRM.gs"]="💼 Сводка AmoCRM - Интеграция AmoCRM"
    ["16_Анализ_Отказов.gs"]="❌ Анализ Отказов - Анализ причин отказов"
)

echo "📂 Создаю готовые файлы в папке ready_modules/"
echo ""

counter=1
total_lines=0

for file in $(ls *_*.gs | sort -V); do
    if [ -f "$file" ]; then
        lines=$(wc -l < "$file")
        total_lines=$((total_lines + lines))
        
        echo "[$counter/16] ${modules[$file]}"
        echo "           Файл: $file"
        echo "           Размер: $lines строк"
        
        # Копируем файл в готовую папку с инструкцией
        cat > "ready_modules/${file}" << EOF
/**
 * ${modules[$file]}
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: ${file%.*} (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

EOF
        
        # Добавляем основной код
        cat "$file" >> "ready_modules/${file}"
        
        echo "           ✅ Готов: ready_modules/$file"
        echo ""
        
        counter=$((counter + 1))
    fi
done

echo "🎉 ПОДГОТОВКА ЗАВЕРШЕНА!"
echo "========================"
echo "📊 Статистика:"
echo "   • Модулей: 16"
echo "   • Общий размер: $total_lines строк"
echo "   • Готовые файлы: ready_modules/"
echo ""
echo "🚀 СЛЕДУЮЩИЕ ШАГИ:"
echo "1. Откройте: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit"
echo "2. Для каждого файла в ready_modules/:"
echo "   • Создайте новый Script file в Google Apps Script"
echo "   • Назовите согласно инструкции в файле"
echo "   • Скопируйте весь код из готового файла"
echo "3. После загрузки всех модулей запустите initializeSystem()"
echo ""
echo "💡 Совет: Откройте файл ready_modules/01_Главный_Скрипт.gs и начните с него!"

# Создаем индексный файл со всеми командами
cat > ready_modules/README.md << EOF
# 🚀 Готовые Модули для Google Apps Script

## Проект ID
\`1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH\`

## Ссылка на проект
https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit

## Порядок установки модулей

EOF

for file in $(ls *_*.gs | sort -V); do
    if [ -f "$file" ]; then
        echo "### ${file%.*}" >> ready_modules/README.md
        echo "- Файл: \`ready_modules/$file\`" >> ready_modules/README.md
        echo "- Описание: ${modules[$file]}" >> ready_modules/README.md
        echo "- Команда копирования: \`cat ready_modules/$file\`" >> ready_modules/README.md
        echo "" >> ready_modules/README.md
    fi
done

cat >> ready_modules/README.md << EOF

## Инструкция по установке

1. Откройте Google Apps Script проект по ссылке выше
2. Для каждого модуля:
   - Создайте новый Script file (+ → Script file)
   - Назовите файл согласно имени модуля (без .gs)
   - Скопируйте код из соответствующего готового файла
   - Вставьте в Google Apps Script
   - Сохраните (Ctrl+S)
3. После установки всех модулей запустите функцию \`initializeSystem()\`

## Общая статистика
- **Модулей**: 16
- **Общий размер**: $total_lines строк кода
- **Язык**: Google Apps Script (JavaScript)
EOF

echo "📄 Создан файл: ready_modules/README.md с полной инструкцией"
