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
    // Inicia o navegador com interface visível (headless: false)
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();

    try {
        // Define um User-Agent para simular um navegador real
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
        await page.setViewport({ width: 1366, height: 768 }); // Define a resolução da página

        // Acessa a página de login
        await page.goto('https://lumma.azurewebsites.net/Security/Login', { waitUntil: 'networkidle2' });
        await page.type('input[type="email"]', 'daniel.aguilar@lumma.com.br'); // Preenche o campo de e-mail
        await page.type('input[type="password"]', 'Bugscp99!'); // Preenche o campo de senha

        // Encontra o botão de login e executa o clique, aguardando a navegação completa
        const loginButton = await page.$('button.btn.btn-lg.btn-primary');
        await Promise.all([
            loginButton.click(),
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
        ]);
        console.log('Login realizado com sucesso!');

        // Acessa o dashboard após o login
        await page.goto('https://lumma.azurewebsites.net/DashBoard/Index', { waitUntil: 'networkidle2' });
        console.log('Navegando para o Dashboard!');

        // Determina o caminho da planilha dentro da pasta 'ERP'
        const userName = os.userInfo().username;
        const folderPath = path.join('C:', 'Users', userName, 'Documents', 'ERP');
        const filePath = path.join(folderPath, 'Automacao.xlsx');

        // Abre a pasta do Explorador de Arquivos
        console.log(`Abrindo o Explorador de Arquivos em: ${folderPath}`);
        exec(`explorer "${folderPath}"`, (err) => {
            if (err) console.error('Erro ao abrir o Explorador:', err);
            else console.log('Explorador de Arquivos aberto com sucesso!');
        });

        await new Promise(resolve => setTimeout(resolve, 2000)); // Aguarda 2 segundos

        // Abre a planilha do Excel
        console.log(`Abrindo o arquivo: ${filePath}`);
        const excel = new winax.Object("Excel.Application");
        excel.Visible = true;
        const workbook = excel.Workbooks.Open(filePath);
        const sheet = workbook.ActiveSheet;

        let row = 2; // Inicia na linha 2 da planilha
        while (true) {
            const cell = sheet.Cells(row, 2);
            if (!cell.Value) {
                console.log(`Fim da planilha na linha ${row}, encerrando...`);
                break;
            }

            console.log(`Processando célula B${row}: ${cell.Value}`);
            await page.bringToFront(); // Traz a página para frente
            await page.click('input#gridOS_searchbar', { clickCount: 3 }); // Seleciona e limpa o campo
            await page.keyboard.press('Backspace');
            await page.type('input#gridOS_searchbar', String(cell.Value)); // Digita o valor
            await page.keyboard.press('Enter');
            console.log('Tecla Enter pressionada!');

            await new Promise(resolve => setTimeout(resolve, 3000)); // Aguarda 3 segundos para carregar os resultados

            // Verifica se a O.S. existe na página
            const osFound = await page.evaluate((cellValue) => {
                const osElements = document.querySelectorAll('td[data-cell="ID"]');
                for (let osElement of osElements) {
                    if (osElement.textContent.trim() === cellValue) {
                        return true;
                    }
                }
                return false;
            }, String(cell.Value));

            if (!osFound) {
                console.log(`O.S. não encontrada para B${row}, pintando a célula de amarelo.`);
                sheet.Cells(row, 2).Interior.Color = 65535; // Define cor amarela na célula
                row++;
                continue;
            }

            console.log('O.S. correspondente encontrada!');

            // Tenta clicar no botão verde se estiver disponível
            const buttonFoundGreen = await page.evaluate(() => {
                const icons = document.querySelectorAll('i.fa.fa-check-circle-o');
                for (let icon of icons) {
                    const color = window.getComputedStyle(icon).color;
                    if (color === 'rgb(0, 166, 90)') {
                        icon.click();
                        return true;
                    }
                }
                return false;
            });

            if (!buttonFoundGreen) {
                console.log('Nenhum botão verde encontrado!');
            } else {
                console.log('Botão verde pressionado!');
                await new Promise(resolve => setTimeout(resolve, 2000));
            }

            await page.keyboard.press('Enter'); // Pressiona Enter para confirmar
            console.log('Tecla Enter pressionada para confirmar ação.');

            // Aguarda a aparição do campo de texto para preenchimento
            const textArea = await page.waitForSelector('textarea.swal2-textarea', { visible: true });
            await textArea.type("O.S já finalizada - Finalizado via automação");
            console.log('Mensagem digitada!');

            // Encontra e clica no botão de envio da mensagem
            const sendButton = await page.waitForSelector('button.swal2-confirm.custom-confirm-button-class', { visible: true });
            await sendButton.click();
            console.log('Botão "Enviar!" clicado!');

            // Aguarda o campo ID reaparecer antes da próxima interação
            await page.waitForSelector('td[data-cell="ID"]', { visible: true });
            console.log('Campo ID reapareceu! Aguardando para prosseguir com a próxima consulta.');

            // Clica no campo de pesquisa novamente e limpa antes de continuar
            await new Promise(resolve => setTimeout(resolve, 10000)); // Fica um pouco mais lento, mas garante que tudo vai carregar certinho pra próxima pesquisa.
            await page.click('input#gridOS_searchbar', { clickCount: 3 });

            row++;
        }

        console.log('Processo concluído com sucesso!');
    } catch (error) {
        console.error('Erro durante a automação:', error);
    } finally {
        // browser.close(); // Manter comentado caso precise visualizar a saída
    }
})();