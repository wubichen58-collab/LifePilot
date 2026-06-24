@dp.message(F.document)
async def handle_incoming_document(message: Message):
    """ИДЕЯ №1: Исправленный прием файлов (.txt и .docx) для RAG-базы знаний"""
    user_id = message.from_user.id
    init_db()
    
    file_name = message.document.file_name
    file_ext = file_name.split(".")[-1].lower() if "." in file_name else ""
    
    # Создаем временное имя для скачивания бинарных файлов
    local_tmp_path = f"tmp_{user_id}_{file_name}"
    
    try:
        if file_ext == "txt" or message.document.mime_type in ["text/plain", "application/octet-stream"]:
            file = await bot.get_file(message.document.file_id)
            file_io = await bot.download_file(file.file_path)
            content = file_io.read().decode('utf-8', errors='ignore')
            
            add_to_knowledge_db(user_id, file_name, content)
            await message.answer(f"✅ Сэр, текстовый документ `{file_name}` успешно отсканирован и занесен в когнитивную матрицу (RAG).")
            
        elif file_ext == "docx":
            # Скачиваем файл на диск, чтобы библиотека docx могла его открыть
            file = await bot.get_file(message.document.file_id)
            await bot.download_file(file.file_path, local_tmp_path)
            
            # Читаем структуру Word документа
            doc = docx.Document(local_tmp_path)
            full_text = []
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    full_text.append(paragraph.text)
                    
            content = "\n".join(full_text)
            
            if not content.strip():
                await message.answer(f"⚠️ Сэр, файл `{file_name}` прочитан, но он пуст или содержит только изображения.")
            else:
                add_to_knowledge_db(user_id, file_name, content)
                await message.answer(f"✅ Протокол RAG: Документ Word `{file_name}` успешно изучен. Данные ({len(content)} симв.) интегрированы в оперативную память.")
                
        else:
            await message.answer(f"❌ Извините, Сэр. Формат `.{file_ext}` пока не поддерживается. Пожалуйста, отправляйте материалы в формате .txt или .docx.")
            
    except Exception as e:
        logger.error(f"Error reading document: {e}")
        await message.answer(f"💥 Произошел сбой при чтении файла `{file_name}`. Мой дешифратор поврежден.")
        
    finally:
        # Чистим временный файл, если он создался
        if os.path.exists(local_tmp_path):
            try: os.remove(local_tmp_path)
            except: pass
