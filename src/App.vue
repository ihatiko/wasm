<template>
  <div class="container">
    <div class="header">
      <h1>📝 DOCX Editor</h1>
      <p>Редактор документов на базе docx-wasm и Vue.js</p>
    </div>

    <div v-if="statusMessage" :class="['status-message', statusMessageType]">
      {{ statusMessage }}
    </div>

    <div class="controls">
      <button @click="addSection" class="btn btn-primary">
        ➕ Добавить секцию
      </button>
      <button @click="saveDocument" class="btn btn-success" :disabled="sections.length === 0 || isLoading">
        💾 Сохранить документ
      </button>
      <button @click="clearAll" class="btn btn-danger" :disabled="sections.length === 0">
        🗑️ Очистить все
      </button>
    </div>

    <div v-if="isLoading" class="loading">
      ⏳ Генерация документа...
    </div>

    <div v-else-if="sections.length === 0" class="empty-state">
      <div class="empty-state-icon">📄</div>
      <div class="empty-state-text">
        Нет секций. Нажмите "Добавить секцию" чтобы начать работу.
      </div>
    </div>

    <div v-else class="sections-container">
      <div v-for="(section, sectionIndex) in sections" :key="sectionIndex" class="section-card">
        <div class="section-header">
          <div class="section-title">Секция {{ sectionIndex + 1 }}</div>
          <div class="section-actions">
            <button @click="addParagraph(sectionIndex)" class="btn btn-primary btn-small">
              ➕ Параграф
            </button>
            <button @click="removeSection(sectionIndex)" class="btn btn-danger btn-small">
              🗑️ Удалить
            </button>
          </div>
        </div>

        <div v-if="section.paragraphs.length === 0" class="empty-state" style="padding: 20px;">
          <div>Нет параграфов в этой секции</div>
        </div>

        <div v-for="(paragraph, paraIndex) in section.paragraphs" :key="paraIndex" class="paragraph-item">
          <div class="paragraph-item-header">
            <span style="font-weight: 600; color: #667eea;">Параграф {{ paraIndex + 1 }}</span>
            <button @click="removeParagraph(sectionIndex, paraIndex)" class="btn btn-danger btn-small">
              ✕
            </button>
          </div>
          
          <div class="form-group">
            <label>Текст:</label>
            <textarea 
              v-model="paragraph.text" 
              @input="updateParagraph(sectionIndex, paraIndex)"
              placeholder="Введите текст параграфа..."
            ></textarea>
          </div>

          <div class="paragraph-controls">
            <div class="checkbox-group">
              <div class="checkbox-item">
                <input 
                  type="checkbox" 
                  :id="`bold-${sectionIndex}-${paraIndex}`"
                  v-model="paragraph.bold"
                  @change="updateParagraph(sectionIndex, paraIndex)"
                >
                <label :for="`bold-${sectionIndex}-${paraIndex}`">Жирный</label>
              </div>
              <div class="checkbox-item">
                <input 
                  type="checkbox" 
                  :id="`italic-${sectionIndex}-${paraIndex}`"
                  v-model="paragraph.italic"
                  @change="updateParagraph(sectionIndex, paraIndex)"
                >
                <label :for="`italic-${sectionIndex}-${paraIndex}`">Курсив</label>
              </div>
              <div class="checkbox-item">
                <input 
                  type="checkbox" 
                  :id="`underline-${sectionIndex}-${paraIndex}`"
                  v-model="paragraph.underline"
                  @change="updateParagraph(sectionIndex, paraIndex)"
                >
                <label :for="`underline-${sectionIndex}-${paraIndex}`">Подчеркнутый</label>
              </div>
            </div>
            
            <div class="form-group" style="margin-top: 10px;">
              <label>Размер шрифта:</label>
              <input 
                type="number" 
                v-model.number="paragraph.fontSize" 
                @input="updateParagraph(sectionIndex, paraIndex)"
                min="8" 
                max="72" 
                style="width: 100px;"
              >
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, onMounted } from 'vue'
import { saveAs } from 'file-saver'

export default {
  name: 'App',
  setup() {
    const sections = ref([])
    const isLoading = ref(false)
    const statusMessage = ref('')
    const statusMessageType = ref('')
    let docxModule = null

    // Инициализация docx-wasm
    onMounted(async () => {
      try {
        isLoading.value = true
        // Динамический импорт для поддержки webpack
        docxModule = await import('docx-wasm')
        showStatus('Библиотека docx-wasm успешно загружена!', 'success')
      } catch (error) {
        console.error('Ошибка загрузки docx-wasm:', error)
        showStatus('Ошибка загрузки библиотеки docx-wasm', 'error')
      } finally {
        isLoading.value = false
      }
    })

    const showStatus = (message, type = 'success') => {
      statusMessage.value = message
      statusMessageType.value = `status-${type}`
      setTimeout(() => {
        statusMessage.value = ''
      }, 5000)
    }

    const addSection = () => {
      sections.value.push({
        paragraphs: []
      })
      showStatus('Секция добавлена', 'success')
    }

    const removeSection = (index) => {
      sections.value.splice(index, 1)
      showStatus('Секция удалена', 'success')
    }

    const addParagraph = (sectionIndex) => {
      sections.value[sectionIndex].paragraphs.push({
        text: '',
        bold: false,
        italic: false,
        underline: false,
        fontSize: 22
      })
    }

    const removeParagraph = (sectionIndex, paraIndex) => {
      sections.value[sectionIndex].paragraphs.splice(paraIndex, 1)
    }

    const updateParagraph = (sectionIndex, paraIndex) => {
      // Параграф обновляется реактивно через v-model
    }

    const clearAll = () => {
      if (confirm('Вы уверены, что хотите удалить все секции?')) {
        sections.value = []
        showStatus('Все секции удалены', 'success')
      }
    }

    const saveDocument = async () => {
      if (!docxModule) {
        showStatus('Библиотека docx-wasm еще не загружена', 'error')
        return
      }

      if (sections.value.length === 0) {
        showStatus('Добавьте хотя бы одну секцию', 'error')
        return
      }

      try {
        isLoading.value = true
        
        const { Docx, Paragraph, Run } = docxModule
        const docx = new Docx()

        // Добавляем каждую секцию
        sections.value.forEach((section, sectionIndex) => {
          // Добавляем заголовок секции
          docx.addParagraph(
            new Paragraph()
              .addRun(
                new Run()
                  .addText(`Секция ${sectionIndex + 1}`)
                  .bold()
                  .size(28)
              )
          )

          // Добавляем параграфы секции
          section.paragraphs.forEach((para) => {
            if (para.text.trim()) {
              const run = new Run().addText(para.text)
              
              // Применяем стили
              if (para.bold) run.bold()
              if (para.italic) run.italic()
              if (para.underline) run.underline()
              if (para.fontSize) run.size(para.fontSize * 2) // docx использует half-points
              
              docx.addParagraph(new Paragraph().addRun(run))
            }
          })

          // Добавляем разрыв между секциями (кроме последней)
          if (sectionIndex < sections.value.length - 1) {
            docx.addParagraph(new Paragraph().addRun(new Run().addBreak()))
          }
        })

        // Добавляем изображение Vue логотипа в конец документа
        try {
          const { Image } = docxModule
          
          if (!Image) {
            throw new Error('Image API не найден в docxModule')
          }
          
          console.log('Начинаем загрузку изображения...')
          
          // Загружаем локальный файл изображения из public папки
          const imageUrl = '/vue-logo.png'
          const response = await fetch(imageUrl)
          
          if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`)
          }
          
          console.log('Изображение загружено, размер:', response.headers.get('content-length'))
          
          const imageBlob = await response.blob()
          const arrayBuffer = await imageBlob.arrayBuffer()
          const bytes = new Uint8Array(arrayBuffer)
          
          console.log('Размер байтов изображения:', bytes.length)
          console.log('Первые байты изображения:', Array.from(bytes.slice(0, 10)))

          // Проверяем, что это действительно PNG (должен начинаться с PNG signature)
          const pngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
          const isPng = pngSignature.every((byte, index) => bytes[index] === byte)
          console.log('Это PNG файл?', isPng)

          // Создаем изображение и устанавливаем размер
          // В DOCX размеры указываются в EMU (English Metric Units)
          // 1 пиксель = 9525 EMU (при 96 DPI)
          // Для 400x400 пикселей: 400 * 9525 = 3,810,000 EMU
          const pixelsToEmu = 9525
          const widthEmu = 400 * pixelsToEmu
          const heightEmu = 400 * pixelsToEmu
          
          const image = new Image(bytes).size(widthEmu, heightEmu)
          
          console.log('Изображение создано, добавляем в документ...')
          console.log('Параметры изображения:', { 
            width: image.w, 
            height: image.h, 
            dataLength: image.data.length,
            widthEmu,
            heightEmu
          })
          
          // Добавляем текст перед изображением для проверки
          docx.addParagraph(
            new Paragraph()
              .addRun(new Run().addText('Vue.js логотип:'))
          )
          
          // Создаем отдельный параграф только с изображением (без текста)
          const imageRun = new Run().addImage(image)
          const imageParagraph = new Paragraph().addRun(imageRun)
          docx.addParagraph(imageParagraph)
          
          // Добавляем текст после изображения для проверки
          docx.addParagraph(
            new Paragraph()
              .addRun(new Run().addText('Конец документа'))
          )
          
          console.log('Изображение успешно добавлено в документ')
          showStatus('Изображение добавлено в документ', 'success')
        } catch (imageError) {
          console.error('Ошибка при добавлении изображения:', imageError)
          console.error('Детали ошибки:', imageError.stack)
          showStatus(`Предупреждение: изображение не добавлено (${imageError.message})`, 'error')
          // Продолжаем без изображения, если не удалось добавить
        }

        // Собираем документ
        const { buffer } = docx.build()
        
        // Сохраняем файл
        const blob = new Blob([buffer], { 
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
        })
        saveAs(blob, `document-${new Date().toISOString().split('T')[0]}.docx`)
        
        showStatus('Документ успешно сохранен!', 'success')
      } catch (error) {
        console.error('Ошибка при сохранении документа:', error)
        showStatus(`Ошибка при сохранении: ${error.message}`, 'error')
      } finally {
        isLoading.value = false
      }
    }

    return {
      sections,
      isLoading,
      statusMessage,
      statusMessageType,
      addSection,
      removeSection,
      addParagraph,
      removeParagraph,
      updateParagraph,
      clearAll,
      saveDocument
    }
  }
}
</script>

