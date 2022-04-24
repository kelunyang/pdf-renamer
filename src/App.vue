<template>
  <v-app>
    <v-sheet class='d-flex flex-column'>
      <v-dialog
        v-model="welcomeW"
        fullscreen
        hide-overlay
        transition="dialog-bottom-transition"
      >
        <v-card>
          <v-toolbar
            dark
            color="primary"
          >
            <v-btn
              icon
              dark
              @click="welcomeW = false"
            >
              <v-icon>fa-times-circle</v-icon>
            </v-btn>
            <v-toolbar-title>歡迎訊息</v-toolbar-title>
          </v-toolbar>
          <v-card-text class='d-flex flex-column ma-1'>
            <v-alert type="info" icon="fa-triangle-exclamation">看懂了就按左上角的X關閉歡迎訊息</v-alert>
            <div class='text-h6'>使用本工具，請務必注意一件事，你所有要更名的PDF檔務必全部放在一個資料夾下面，不用在裡面分層，程式只會掃第一層而已！</div>
            <v-img src="@/assets/tip.png"></v-img>
            <div class='text-caption'>kelunyang@outlook.com 2022</div>
            <div class='blue--text' @click='openGitHub'>點此打開本程式的GitHub</div>
          </v-card-text>
        </v-card>
      </v-dialog>
      <v-dialog
        v-model="viewRuleW"
        fullscreen
        hide-overlay
        transition="dialog-bottom-transition"
      >
        <v-card>
          <v-toolbar
            dark
            color="primary"
          >
            <v-btn
              icon
              dark
              @click="viewRuleW = false"
            >
              <v-icon>fa-times-circle</v-icon>
            </v-btn>
            <v-toolbar-title>PDF更名規則：[{{ selectedRule.type ? "內容" : "中介" }}]{{ selectedRule.keyword }}</v-toolbar-title>
          </v-toolbar>
          <v-card-text class='d-flex flex-column ma-1'>
            <div class='text-h6'>重複次數：{{ selectedRule.times }}</div>
            <div class='text-h6'>更名目標：{{ selectedRule.result }}</div>
            <v-simple-table v-if='selectedRule.occurrence.length > 0'>
              <thead>
                <tr>
                  <th
                    class="text-left"
                  >
                    檔案名稱
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr
                  v-for="pdf in selectedRule.occurrence"
                  :key="pdf.id + 'occ'"
                >
                  <td class="text-left">
                    {{ pdf.name }}
                  </td>
                </tr>
              </tbody>
            </v-simple-table>
          </v-card-text>
        </v-card>
      </v-dialog>
      <v-dialog
        v-model="renameRulesW"
        fullscreen
        hide-overlay
        transition="dialog-bottom-transition"
      >
        <v-card>
          <v-toolbar
            dark
            color="primary"
          >
            <v-btn
              icon
              dark
              @click="renameRulesW = false"
            >
              <v-icon>fa-times-circle</v-icon>
            </v-btn>
            <v-toolbar-title>PDF更名規則清單</v-toolbar-title>
          </v-toolbar>
          <v-card-text class='d-flex flex-column ma-1'>
            <v-alert type="error" icon="fa-skull" v-if='ruleError !== ""'>{{ ruleError }}</v-alert>
            <v-alert type="info" icon="fa-info" v-if='renameRules.length > 0'>已讀入了{{ renameRules.length }}條規則</v-alert>
            <v-alert type="info" icon="fa-info">
              <v-row>
                <v-col class='grow'>
                  程式會逐個打開PDF，逐頁掃描其內文（或PDF中介資料），比對你給的關鍵字之後更名，請務必給足以辨識的關鍵字（如身分證字號），否則可能會不精準又浪費時間（關鍵字可以是正規表達式）
                </v-col>
                <v-col class='shrink'>
                  <v-btn @click='downloadCSV(sampleRule,"範例更名原則")'>按此下載範例</v-btn>
                </v-col>
              </v-row>
            </v-alert>
            <v-file-input accept="text/csv" prepend-icon='fa-file-csv' outlined v-model="ruleFile" placeholder="請選擇你要讀入的PDF規則檔"/>
            <v-simple-table v-if='renameRules.length > 0'>
              <thead>
                <tr>
                  <th
                    class="text-center"
                  >
                    原則
                  </th>
                  <th
                    class="text-left"
                  >
                    搜尋關鍵字
                  </th>
                  <th
                    class="text-left"
                  >
                    結果名稱
                  </th>
                  <th
                    class="text-left"
                  >
                    重複次數
                  </th>
                  <th
                    class="text-right"
                  >
                    搜尋結果
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr
                  v-for="rule in renameRules"
                  :key="rule.id"
                >
                  <td class="text-center">
                    {{ rule.type ? "內容" : "中介" }}
                  </td>
                  <td class="text-left">
                    {{ rule.keyword }}
                  </td>
                  <td class="text-left">
                    {{ rule.result }}
                  </td>
                  <td class="text-left">
                    {{ rule.times }}
                  </td>
                  <td class="text-right">
                    {{ rule.occurrence.length }}
                    <v-btn
                      v-if='rule.occurrence.length > 0'
                      icon
                      @click="viewRule(rule)"
                    >
                      <v-icon>fa-magnifying-glass</v-icon>
                    </v-btn>
                  </td>
                </tr>
              </tbody>
            </v-simple-table>
          </v-card-text>
        </v-card>
      </v-dialog>
      <v-dialog
        v-model="pdfFilesW"
        fullscreen
        hide-overlay
        transition="dialog-bottom-transition"
      >
        <v-card>
          <v-toolbar
            dark
            color="primary"
          >
            <v-btn
              icon
              dark
              @click="pdfFilesW = false"
            >
              <v-icon>fa-times-circle</v-icon>
            </v-btn>
            <v-toolbar-title>PDF檔案現在存放位置</v-toolbar-title>
          </v-toolbar>
          <v-card-text class='d-flex flex-column ma-1'>
            <v-alert type="info" icon="fa-info" v-if='pdfFiles.length > 0'>已找到了{{ pdfFiles.length }}個PDF</v-alert>
            <v-alert type="info" icon="fa-info">請注意，程式之後會在這個資料夾建立一個「export」的資料夾之後輸出更名的PDF</v-alert>
            <v-file-input prepend-icon='fa-folder' webkitdirectory directory multiple outlined v-model="selectedFolder" placeholder="請選擇PDF檔案現在存放的資料夾"/>
          </v-card-text>
        </v-card>
      </v-dialog>
      <v-row class='mb-1 pa-2'>
        <v-col class='blue darken-4 justify-space-between flex-row align-content-space-between'>
          <div class='text-caption white--text ma-1 text-left'>更名規則</div>
          <div class='text-h4 white--text ma-1 text-center'>{{ renameRules.length }}</div>
          <v-btn @click="renameRulesW = true" class='red darken-4 white--text ma-1'>{{ renameRules.length > 0 ? '重建' : '讀入' }}更名規則</v-btn>
        </v-col>
        <v-col class='blue darken-4 justify-space-between flex-row align-content-space-between'>
          <div class='text-caption white--text ma-1 text-left'>讀入PDF檔案數</div>
          <div class='text-h4 white--text text-center'>{{ pdfFiles.length }}</div>
          <v-btn @click="pdfFilesW = true" class='red darken-4 white--text ma-1'>設定PDF所在的資料夾</v-btn>
        </v-col>
      </v-row>
      <v-alert type="info" icon="fa-info" v-if='currentMsg !== ""'>{{ currentMsg }}</v-alert>
      <v-slider
        :label='"同時掃描" + scanQueue + "個PDF"'
        min='1'
        max='5'
        v-model="scanQueue"
        thumb-label
        v-if='paraSetted'
      ></v-slider>
      <v-btn @click='scanPDF(false)' class='indigo darken-4 white--text ma-1' v-if='paraSetted'>掃描文件</v-btn>
      <v-btn @click='renamePDF(false)' class='indigo darken-4 white--text ma-1' v-if='allScaned'>文件更名</v-btn>
      <v-simple-table v-if='pdfFiles.length > 0'>
        <thead>
          <tr>
            <th
              class="text-left"
            >
              檔名
            </th>
            <th
              class="text-right"
            >
              掃描狀態
            </th>
            <th
              class="text-right"
            >
              符合規則
            </th>
            <th
              class="text-right"
            >
              更名
            </th>
          </tr>
        </thead>
        <tbody>
          <tr
            v-for="pdf in pdfFiles"
            :key="pdf.id"
            :class="pdf.rule === undefined ? 'red lighten-5' : ''"
          >
            <td class="text-left">
              {{ pdf.name }}
            </td>
            <td class='text-right'>
              {{ pdf.current ? "正在掃描" : pdf.scaned ? "已掃描" : "未掃描" }}
            </td>
            <td class='text-right'>
              {{ pdf.rule === undefined ? "否" : "是" }}
              <v-btn
                v-if='pdf.rule !== undefined'
                icon
                @click="viewRule(pdf.rule)"
              >
                <v-icon>fa-magnifying-glass</v-icon>
              </v-btn>
            </td>
            <td class='text-right'>
              {{ pdf.exported ? "是" : "否" }}
            </td>
          </tr>
        </tbody>
      </v-simple-table>
    </v-sheet>
  </v-app>
</template>

<style>
@import url('https://fonts.googleapis.com/css?family=Noto+Sans+TC:100,300,400,500,700,900&display=swap');
html {
  scroll-behavior: smooth;
}
#app {
  font-family: 'Noto Sans TC', sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}
h1 {
  font-family: 'Noto Sans TC', sans-serif;
  font-weight: 900;
}
</style>

<script>
import _ from 'lodash';
import { v4 as uuidv4 } from 'uuid';
import Papa from 'papaparse';
import '@fortawesome/fontawesome-free/css/all.css';
import '@fortawesome/fontawesome-free/js/all.js';
import iconv from 'iconv-lite';
import chardet from 'chardet';
import dayjs from 'dayjs';
import { Buffer } from 'buffer';
let duration = require('dayjs/plugin/duration');
dayjs.extend(duration);

export default {
  name: 'pdfRenamer',
  mounted: function() {
    let oriobj = this;
    window.ipcRenderer.receive('scanFolder', (args) => {
      oriobj.pdfFiles = [];
      oriobj.importFolder = args.folder;
      for(let i=0; i<args.pdfs.length; i++) {
        oriobj.pdfFiles.push({
          id: uuidv4(),
          name: args.pdfs[i].name,
          fullpath: args.pdfs[i].fullpath,
          exported: false,
          errorReason: '',
          scaned: false,
          current: false,
          rule: undefined,
        })
      }
    });
    window.ipcRenderer.receive('scanStatus', (args) => {
      oriobj.currentMsg = args;
    });
    window.ipcRenderer.receive('scanPDF', (args) => {
      let workingPDF = _.filter(oriobj.pdfFiles, (pdf) => {
        return pdf.id === args.pdf.id;
      });
      if(workingPDF.length > 0) {
        workingPDF[0].current = false;
        workingPDF[0].scaned = true;
        workingPDF[0].rule = args.rule;
      }
      if(args.rule !== undefined) {
        let workingRule = _.filter(oriobj.renameRules, (rule) => {
          return rule.id === args.rule.id
        });
        if(workingRule.length > 0) {
          workingRule[0].occurrence.push(workingPDF[0]);
        }
      }
      let queue = _.filter(oriobj.pdfFiles, (pdf) => {
        return !pdf.scaned;
      })
      if(queue.length > 0) {
        oriobj.scanPDF(true);
      } else {
        let working = _.filter(oriobj.pdfFiles, (pdf) => {
          return pdf.current;
        });
        if(working.length === 0) {
          oriobj.currentMsg = "掃描共花了：" + dayjs.duration((dayjs().valueOf() - oriobj.startScan)).format('HH:mm:ss:SSS');
          setTimeout(() => {
            oriobj.currentMsg = "";
          }, 3000);
        }
      }
    });
    window.ipcRenderer.receive('renamePDF', (args) => {
      let workingPDF = _.filter(oriobj.pdfFiles, (pdf) => {
        return pdf.id === args.id;
      });
      if(workingPDF.length > 0) {
        workingPDF[0].current = false;
        workingPDF[0].exported = true;
      }
      let queue = _.filter(oriobj.pdfFiles, (pdf) => {
        if(pdf.rule !== undefined) {
          return !pdf.exported;
        }
        return false;
      })
      if(queue.length > 0) {
        oriobj.renamePDF(true);
      } else {
        let working = _.filter(oriobj.pdfFiles, (pdf) => {
          return pdf.current;
        });
        if(working.length === 0) {
          oriobj.currentMsg = "輸出共花了：" + dayjs.duration(dayjs().valueOf() - oriobj.startRename).format('HH:mm:ss:SSS');
          setTimeout(() => {
            oriobj.currentMsg = "";
          }, 3000);
        }
      }
    });
  },
  methods: {
    openGitHub: function() {
      window.ipcRenderer.send("openLink", "https://github.com/kelunyang/pdf-renamer");
    },
    viewRule: function(rule) {
      let selectRule = _.filter(this.renameRules, (rrule) => {
        return rule.id === rrule.id;
      });
      if(selectRule.length > 0) {
        this.selectedRule = selectRule[0];
        this.viewRuleW = true;
      }
    },
    scanPDF: function(rec) {
      if(!rec) { this.startScan = dayjs().valueOf(); }
      let unScaned = _.filter(this.pdfFiles, (file) => {
        if(!file.current) {
          return !file.scaned;
        }
        return false;
      });
      let scanning = _.filter(this.pdfFiles, (file) => {
        return file.current;
      });
      let queueNum = this.scanQueue < unScaned.length ? this.scanQueue - scanning.length : unScaned.length;
      if(queueNum > 0) {
        let queue = _.slice(_.shuffle(unScaned), 0, queueNum);
        for(let i=0; i<queue.length; i++) {
          queue[i].current = true;
          window.ipcRenderer.send("scanPDF", {
            pdf: queue[i],
            rules: this.renameRules
          });
        }
      }
    },
    renamePDF: function(rec) {
      if(!rec) { this.startRename = dayjs().valueOf(); }
      let unExported = _.filter(this.pdfFiles, (file) => {
        if(file.rule !== undefined) {
          if(!file.current) {
            return !file.exported;
          }
        }
        return false;
      });
      let scanning = _.filter(this.pdfFiles, (file) => {
        return file.current;
      });
      let queueNum = this.scanQueue < unExported.length ? this.scanQueue - scanning.length : unExported.length;
      if(queueNum > 0) {
        let queue = _.slice(_.shuffle(unExported), 0, queueNum);
        for(let i=0; i<queue.length; i++) {
          queue[i].current = true;
          window.ipcRenderer.send("renamePDF", {
            pdf: queue[i],
            folder: this.importFolder
          });
        }
      }
    },
    downloadCSV: function(arr, filename) {
      let output = "\ufeff"+ Papa.unparse(arr);
      let element = document.createElement('a');
      let blob = new Blob([output], { type: 'text/csv' });
      let url = window.URL.createObjectURL(blob);
      element.setAttribute('href', url);
      element.setAttribute('download', filename + ".csv");
      element.click();
    },
  },
  computed: {
    paraSetted: function() {
      if(this.importFolder !== "") {
        if(this.renameRules.length > 0) {
          return true;
        }
      }
      return false;
    },
    allScaned: function() {
      if(this.paraSetted) {
        let unScaned = _.filter(this.pdfFiles, (pdf) => {
          return !pdf.scaned;
        });
        return unScaned.length === 0;
      }
      return false;
    },
    exportedPDF: function() {
      return _.filter(this.pdfFiles, (pdf) => {
        return pdf.exported
      });
    }
  },
  data: () => ({
    welcomeW: true,
    startScan: 0,
    startRename: 0,
    selectedRule: {
      keyword: "",
      occurrence: [],
      times: 0,
      result: "",
      type: true
    },
    viewRuleW: false,
    scanQueue: 1,
    currentMsg: "",
    pdfFilesW: false,
    renameRulesW: false,
    renameRules: [],
    pdfFiles: [],
    ruleFile: undefined,
    ruleError: '',
    selectedFolder: [],
    importFolder: "",
    sampleRule: [
      {
        關鍵字: '史黛拉',
        目標檔名: '小熊維尼',
        原則: '內容',
        重複次數: 1
      },
      {
        關鍵字: '黛西',
        目標檔名: '唐老鴨',
        原則: '中介',
        重複次數: 3
      }
    ]
  }),
  watch: {
    selectedFolder: async function() {
      if(this.selectedFolder.length > 0) {
        window.ipcRenderer.send("scanFolder", this.selectedFolder[0].path);
      }
    },
    ruleFile: {
      immediate: true,
      async handler () {
        let oriobj = this;
        if (this.ruleFile !== undefined) {
          let reader = new FileReader();
          reader.readAsArrayBuffer(oriobj.ruleFile);
          reader.onload = ((file) => {
            try {
              let result = Buffer.from(file.target.result);
              let encoding = chardet.detect(result);
              let content = iconv.decode(Buffer.from(file.target.result), encoding);
              oriobj.ruleError = '';
              Papa.parse(content, {
                header: true,
                skipEmptyLines: true,
                complete: async function(result) {
                  if(result.data.length > 0) {
                    for(let i=0; i<result.data.length; i++) {
                      if(result.data[i]['原則'].match(/內容|中介/g)) {
                        oriobj.renameRules.push({
                          id: uuidv4(),
                          keyword: result.data[i]['關鍵字'],
                          result: result.data[i]['目標檔名'],
                          type: result.data[i]['原則'] === '內容',
                          times: parseInt(result.data[i]['重複次數']),
                          occurrence: []
                        });
                      } else {
                        oriobj.ruleError = result.data[i]['關鍵字'] + '原則設定錯誤（打錯字？）';
                      }
                    }
                  }
                }
              });
            } catch(e) {
              console.dir(e);
            }
          });
        }
      }
    },
  }
};
</script>
