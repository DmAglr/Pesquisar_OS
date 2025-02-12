/**
 * Automação utilizando Puppeteer para acessar o sistema Lumma, realizar login,
 * buscar informações em uma planilha do Excel e interagir com a interface da aplicação web.
 * ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 * @author Daniel Aguilar
 * @since 11/02/2025
 */

// Importação dos módulos necessários
const puppeteer = require('puppeteer'); // Controle do navegador
const { exec } = require('child_process'); // Execução de comandos no sistema operacional
const os = require('os'); // Informações do sistema operacional
const path = require('path'); // Manipulação de caminhos de arquivos
const winax = require('winax'); // Controle do Excel via COM API (Windows)

(async () => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();

    try {
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
        await page.setViewport({ width: 1366, height: 768 });

        await page.goto('https://lumma.azurewebsites.net/Security/Login', { waitUntil: 'networkidle2' });
        await page.type('input[type="email"]', 'daniel.aguilar@lumma.com.br');
        await page.type('input[type="password"]', 'Bugscp99!');

        const loginButton = await page.$('button.btn.btn-lg.btn-primary');
        await Promise.all([
            loginButton.click(),
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
        ]);
        console.log('Login realizado com sucesso!');

        await page.goto('https://lumma.azurewebsites.net/DashBoard/Index', { waitUntil: 'networkidle2' });
        console.log('Navegando para o Dashboard!');

        const userName = os.userInfo().username;
        const folderPath = path.join('C:', 'Users', userName, 'Documents', 'ERP');
        const filePath = path.join(folderPath, 'Automacao.xlsx');

        console.log(`Abrindo o Explorador de Arquivos em: ${folderPath}`);
        exec(`explorer "${folderPath}"`, (err) => {
            if (err) console.error('Erro ao abrir o Explorador:', err);
            else console.log('Explorador de Arquivos aberto com sucesso!');
        });

        await new Promise(resolve => setTimeout(resolve, 2000));

        console.log(`Abrindo o arquivo: ${filePath}`);
        const excel = new winax.Object("Excel.Application");
        excel.Visible = true;
        const workbook = excel.Workbooks.Open(filePath);
        const sheet = workbook.ActiveSheet;

        let row = 2;
        while (true) {
            const cell = sheet.Cells(row, 2);
            if (!cell.Value) {
                console.log(`Fim da planilha na linha ${row}, encerrando...`);
                break;
            }

            console.log(`Processando célula B${row}: ${cell.Value}`);
            await page.bringToFront();
            await page.click('input#gridOS_searchbar', { clickCount: 3 });
            await page.keyboard.press('Backspace');
            await page.type('input#gridOS_searchbar', String(cell.Value));
            await page.keyboard.press('Enter');
            console.log('Tecla Enter pressionada!');

            await new Promise(resolve => setTimeout(resolve, 3000));

            const osElements = await page.$$('div.e-gridcontent td[data-cell="ID"]');
            let matchingRow = null;

            for (const osElement of osElements) {
                const text = await page.evaluate(el => el.textContent.trim(), osElement);
                if (text === String(cell.Value)) {
                    matchingRow = osElement;
                    break;
                }
            }

            if (!matchingRow) {
                console.log(`O.S. não encontrada para B${row}, pintando a célula de amarelo.`);
                sheet.Cells(row, 2).Interior.Color = 65535;
                row++;
                continue;
            }

            console.log('O.S. correspondente encontrada!');

            const parentRow = await page.evaluateHandle(el => el.closest('tr'), matchingRow);
            const buttonGreen = await parentRow.$('i.fa.fa-check-circle-o');
            
            if (buttonGreen) {
                const color = await page.evaluate(el => getComputedStyle(el).color, buttonGreen);
                if (color === 'rgb(0, 166, 90)') {
                    await buttonGreen.click();
                    console.log('Botão verde pressionado!');
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            } else {
                console.log('Nenhum botão verde encontrado na linha correspondente!');
            }

            await page.keyboard.press('Enter');
            console.log('Tecla Enter pressionada para confirmar ação.');

            const textArea = await page.waitForSelector('textarea.swal2-textarea', { visible: true });
            await textArea.type("O.S já finalizada - Finalizado via automação");
            console.log('Mensagem digitada!');

            const sendButton = await page.waitForSelector('button.swal2-confirm.custom-confirm-button-class', { visible: true });
            await sendButton.click();
            console.log('Botão "Enviar!" clicado!');

            await page.waitForSelector('td[data-cell="ID"]', { visible: true });
            console.log('Campo ID reapareceu! Aguardando para prosseguir com a próxima consulta.');

            await new Promise(resolve => setTimeout(resolve, 10000));
            await page.click('input#gridOS_searchbar', { clickCount: 3 });

            row++;
        }

        console.log('Processo concluído com sucesso!');
    } catch (error) {
        console.error('Erro durante a automação:', error);
    } finally {
        // browser.close();
    }
})();
