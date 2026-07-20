import type { App } from 'obsidian'
import type { UniverPluginSettings } from '@/types/setting'
import { Modal } from 'obsidian'
import { uiText } from '@/i18n'
import { createNewFile } from '@/utils/file'

interface ModalText {
  title: string
  docBtn: string
  sheetBtn: string
  excelBtn: string
  wordBtn: string
}

export class ChooseTypeModal extends Modal {
  settings: UniverPluginSettings
  constructor(app: App, settings: UniverPluginSettings) {
    super(app)
    this.settings = settings
  }

  onOpen(): void {
    const { contentEl } = this
    this.titleEl.setText(this.getModalText().title)

    const btnContainer = contentEl.createDiv()
    btnContainer.classList.add('univer-modal-btn-container')

    const docBtn = btnContainer.createEl('button', {
      text: this.getModalText().docBtn,
      cls: 'univer-modal-btn',
    })

    const sheetBtn = btnContainer.createEl('button', {
      text: this.getModalText().sheetBtn,
      cls: 'univer-modal-btn',
    })

    const excelBtn = this.settings.isSupportXlsx
      ? btnContainer.createEl('button', {
          text: this.getModalText().excelBtn,
          cls: 'univer-modal-btn',
        })
      : undefined

    const wordBtn = this.settings.isSupportDocx
      ? btnContainer.createEl('button', {
          text: this.getModalText().wordBtn,
          cls: 'univer-modal-btn',
        })
      : undefined

    docBtn.onclick = () => {
      void createNewFile(this.app, 'udoc', this.settings.language)
      this.close()
    }

    sheetBtn.onclick = () => {
      void createNewFile(this.app, 'usheet', this.settings.language)
      this.close()
    }

    if (excelBtn) {
      excelBtn.onclick = () => {
        void createNewFile(this.app, 'xlsx', this.settings.language)
        this.close()
      }
    }

    if (wordBtn) {
      wordBtn.onclick = () => {
        void createNewFile(this.app, 'docx', this.settings.language)
        this.close()
      }
    }
  }

  onClose(): void {
  // 清空 contentEl 元素的内容
    this.contentEl.empty()
  }

  getModalText(): ModalText {
    const text = uiText(this.settings.language)
    return {
      title: text.chooseType,
      docBtn: text.univerDoc,
      sheetBtn: text.univerSheet,
      excelBtn: text.newExcel,
      wordBtn: text.newWord,
    }
  }
}
