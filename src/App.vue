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
      <label class="btn btn-primary file-upload-btn">
        📤 Загрузить DOCX файл
        <input 
          type="file" 
          accept=".docx" 
          @change="handleFileUpload" 
          style="display: none;"
          :disabled="isLoading"
        >
      </label>
      <button @click="addStampToDocument" class="btn btn-success" :disabled="!uploadedDocx || isLoading">
        🔖 Поставить печать
      </button>
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

    <div v-if="uploadedFileName" class="uploaded-file-info">
      📄 Загружен файл: <strong>{{ uploadedFileName }}</strong>
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
import JSZip from 'jszip'

export default {
  name: 'App',
  setup() {
    const sections = ref([])
    const isLoading = ref(false)
    const statusMessage = ref('')
    const statusMessageType = ref('')
    const uploadedDocx = ref(null)
    const uploadedFileName = ref('')
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

    // Загрузка DOCX файла
    const handleFileUpload = async (event) => {
      const file = event.target.files[0]
      if (!file) return

      if (!file.name.endsWith('.docx')) {
        showStatus('Пожалуйста, выберите файл .docx', 'error')
        return
      }

      try {
        isLoading.value = true
        const arrayBuffer = await file.arrayBuffer()
        uploadedDocx.value = arrayBuffer
        uploadedFileName.value = file.name
        showStatus(`Файл "${file.name}" успешно загружен!`, 'success')
      } catch (error) {
        console.error('Ошибка загрузки файла:', error)
        showStatus(`Ошибка загрузки файла: ${error.message}`, 'error')
      } finally {
        isLoading.value = false
      }
    }

    // Загрузка изображения печати
    const loadStampImage = async () => {
      try {
        // Загружаем SVG файл печати
        const svgUrl = '/src/test.svg'
        const response = await fetch(svgUrl)

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`)
        }

        const svgText = await response.text()

        // Конвертируем SVG в PNG через canvas
        const canvas = document.createElement('canvas')
        const ctx = canvas.getContext('2d')
        const img = new window.Image()

        const imageBytes = await new Promise((resolve, reject) => {
          img.onload = async () => {
            // Устанавливаем размер canvas для печати (обычно печати небольшие)
            canvas.width = 300
            canvas.height = 300

            // Рисуем SVG на canvas
            ctx.drawImage(img, 0, 0, 300, 300)

            // Конвертируем canvas в PNG blob
            canvas.toBlob(async (blob) => {
              const arrayBuffer = await blob.arrayBuffer()
              resolve(new Uint8Array(arrayBuffer))
            }, 'image/png')
          }

          img.onerror = () => reject(new Error('Не удалось загрузить SVG печати'))

          // Загружаем SVG как data URL
          const svgBlob = new Blob([svgText], { type: 'image/svg+xml' })
          img.src = URL.createObjectURL(svgBlob)
        })

        return imageBytes
      } catch (error) {
        console.error('Ошибка загрузки изображения печати:', error)
        throw error
      }
    }

    // Добавление печати используя только docx-wasm
    const addStampToDocument = async () => {
      if (!uploadedDocx.value || !docxModule) {
        showStatus('Сначала загрузите DOCX файл', 'error')
        return
      }

      try {
        isLoading.value = true
        showStatus('Обработка документа...', 'success')

        const stampImageBytes = await loadStampImage()
        const { Docx, Paragraph, Run, Image } = docxModule
        
        // Создаем документ с печатью через docx-wasm
        const docx = new Docx()
        const pixelsToEmu = 9525
        const stampImage = new Image(stampImageBytes).size(300 * pixelsToEmu, 300 * pixelsToEmu)
        docx.addParagraph(new Paragraph().addRun(new Run().addImage(stampImage)))
        const { buffer: stampBuffer } = docx.build()
        
        // Загружаем архивы
        const originalZip = await JSZip.loadAsync(uploadedDocx.value)
        const stampZip = await JSZip.loadAsync(stampBuffer)
        const newZip = await JSZip.loadAsync(uploadedDocx.value)
        
        // Объединяем document.xml
        const originalDocXml = await originalZip.file('word/document.xml').async('string')
        const stampDocXml = await stampZip.file('word/document.xml').async('string')
        const originalBodyEnd = originalDocXml.lastIndexOf('</w:body>')
        const stampBodyStart = stampDocXml.indexOf('<w:body')
        const stampBodyEnd = stampDocXml.indexOf('</w:body>')
        
        if (originalBodyEnd === -1 || stampBodyStart === -1 || stampBodyEnd === -1) {
          throw new Error('Не удалось найти body в документе')
        }
        
        const stampBodyContent = stampDocXml.substring(
          stampDocXml.indexOf('>', stampBodyStart) + 1,
          stampBodyEnd
        )
        
        newZip.file('word/document.xml', 
          originalDocXml.substring(0, originalBodyEnd) + 
          stampBodyContent + 
          originalDocXml.substring(originalBodyEnd)
        )
        
        // Объединяем relationships
        const originalRelsXml = await originalZip.file('word/_rels/document.xml.rels').async('string')
        const stampRelsXml = await stampZip.file('word/_rels/document.xml.rels').async('string')
        const relsEndIndex = originalRelsXml.lastIndexOf('</Relationships>')
        const stampRelsStart = stampRelsXml.indexOf('<Relationships')
        const stampRelsEnd = stampRelsXml.indexOf('</Relationships>')
        
        if (relsEndIndex !== -1 && stampRelsStart !== -1 && stampRelsEnd !== -1) {
          const stampRelsContent = stampRelsXml.substring(
            stampRelsXml.indexOf('>', stampRelsStart) + 1,
            stampRelsEnd
          )
          newZip.file('word/_rels/document.xml.rels',
            originalRelsXml.substring(0, relsEndIndex) + 
            stampRelsContent + 
            originalRelsXml.substring(relsEndIndex)
          )
        }
        
        // Копируем изображения
        const imageFiles = Object.keys(stampZip.files).filter(
          path => path.startsWith('word/media/') && !stampZip.files[path].dir
        )
        for (const imagePath of imageFiles) {
          const imageFile = stampZip.file(imagePath)
          if (imageFile) {
            newZip.file(imagePath, await imageFile.async('uint8array'))
          }
        }
        
        // Сохраняем файл
        const blob = new Blob([await newZip.generateAsync({ type: 'arraybuffer' })], { 
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
        })
        saveAs(blob, uploadedFileName.value.replace('.docx', '') + '_с_печатью.docx')
        
        showStatus('Печать успешно добавлена! Документ сохранен.', 'success')
      } catch (error) {
        console.error('Ошибка при добавлении печати:', error)
        showStatus(`Ошибка: ${error.message}`, 'error')
      } finally {
        isLoading.value = false
      }
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
      uploadedDocx,
      uploadedFileName,
      addSection,
      removeSection,
      addParagraph,
      removeParagraph,
      updateParagraph,
      clearAll,
      saveDocument,
      handleFileUpload,
      addStampToDocument
    }
  }
}
</script>

