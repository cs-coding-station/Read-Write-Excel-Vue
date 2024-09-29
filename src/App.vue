<template>
  <main>
    <v-container fluid>
      <v-row justify="center">
        <v-col cols="10">
          <v-row dense no-gutters justify="space-between">
            <v-file-input
              v-model="file"
              label="File input"
              variant="solo"
              density="compact"
              @change="readExcelFile"
              class="file"
            ></v-file-input>

            <v-btn color="blue" @click="downloadFile">Download File</v-btn>
          </v-row>
        </v-col>
      </v-row>

      <v-row justify="center">
        <v-col cols="10">
          <v-card>
            <v-data-table :headers="header" :items="items" :height="height - 200"></v-data-table>
          </v-card>
        </v-col>
      </v-row>
    </v-container>
  </main>
</template>

<script setup>
import readXlsxFile from 'read-excel-file'
import writeXlsxFile from 'write-excel-file'

import { ref } from 'vue'
import { useDisplay } from 'vuetify'

const { height } = useDisplay()

const file = ref([])

const header = [
  { key: 'name', title: 'Name' },
  { key: 'address', title: 'Address' },
  { key: 'job', title: 'Job' },
  { key: 'phone', title: 'Phone Number' }
]

const items = ref([])

function readExcelFile() {
  readXlsxFile(file.value[0]).then((rows) => {
    rows.forEach((row, index) => {
      if (index > 0) {
        items.value.push({
          name: row[0],
          address: row[1],
          job: row[2],
          phone: row[3]
        })
      }
    })
  })
}

async function downloadFile() {
  const rows = []
  const header = [
    { value: 'Name', fontWeight: 'bold' },
    { value: 'Address', fontWeight: 'bold' },
    { value: 'Job', fontWeight: 'bold' },
    { value: 'Phone Number', fontWeight: 'bold' }
  ]

  rows.push(header)

  items.value.forEach((item) => {
    let row = []
    row.push(
      {
        type: String,
        value: item.name
      },
      {
        type: String,
        value: item.address
      },
      {
        type: String,
        value: item.job
      },
      {
        type: Number,
        value: item.phone
      }
    )

    rows.push(row)
  })

  await writeXlsxFile(rows, {
    fileName: 'file.xlsx'
  })
}
</script>

<style scoped>
main {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100vh;
}

.file {
  max-width: 400px;
}
</style>
