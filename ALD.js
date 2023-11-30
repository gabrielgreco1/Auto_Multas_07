import puppeteer from 'puppeteer';
import Jimp from 'Jimp';
import XLSX from 'xlsx';
import sendEmail from '../Modules/Email/Email.js'
import deletePngFiles from '../Modules/Delete_img/Delete_file.js'
import 'dotenv/config';

function paths(){
    const directory1 = 'C:\\Users\\ggreco\\Documents\\Automações\\Code\\Node\\Multas\\PRD\\07 - Multas\\images';
    const directory2 = 'S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno'

    deletePngFiles(directory1);
    deletePngFiles(directory2);
}

paths();

// Ler o arquivo Excel
const workbook = XLSX.readFile('S:\\Automacoes\\Multas\\07 - Multas ALD\\Ald_valida.xlsx');
const sheet_name_list = workbook.SheetNames;
const dados = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

// const workbook_retorno = XLSX.readFile('S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Validacao_retorno.xlsx');
// const sheet_name_list_retorno = workbook_retorno.SheetNames;
// const dados_retorno = XLSX.utils.sheet_to_json(workbook_retorno.Sheets[sheet_name_list_retorno[0]]);

// console.error(dados_retorno)
const email = dados[0]['User Process']

console.log('--------------------------//--------------------------//--------------------------');
console.log('--------------------------//--------------------------//--------------------------');
console.log(`--------------------------//     START AUTOMATION - USER:  ${email}    //-----`);
console.log('--------------------------//--------------------------//--------------------------');
console.log('--------------------------//--------------------------//--------------------------');

(async () => {
  // Iniciar o navegador
    const browser = await puppeteer.launch({ headless: 'new' }); // 'new' para rodar em background
    await new Promise(resolve => setTimeout(resolve, 5000));
    // Abrir uma nova página
    const page = await browser.newPage();

    // Definir a resolução da janela
    await page.setViewport({ width: 1280, height: 1024 });

    // Navegar até o URL especificado
    await page.goto(process.env.process);

    // Inserir login, senha e logar
    await page.type('#TxtLogin', process.env.user);
    await page.type('#TxtPassword', process.env.password);
    await page.keyboard.press('Enter');


    let contador_inicio = 0;
    await new Promise(resolve => setTimeout(resolve, 5000));

    while (contador_inicio < 11){
        const xinicio = 158;
        const yinicio = 401;
        await page.screenshot({ path: 'images/screenshot_inicio.png' });
        const img_inicio = await Jimp.read('images/screenshot_inicio.png');
        const corPixel = Jimp.intToRGBA(img_inicio.getPixelColor(xinicio, yinicio));
        let cor = corPixel.r === 70 && corPixel.g === 83 && corPixel.b === 123;
        if (cor) {
            console.log(`${new Date().toLocaleString()} - Sistema carregou`);
            contador_inicio = 11;
        } else {
            await new Promise(resolve => setTimeout(resolve, 2000));
            contador_inicio = contador_inicio + 1;

            if (contador_inicio > 10) {
                console.log(`${new Date().toLocaleString()} - Não foi possível entrar no sistema`);
                await browser.close();
            }
    }}

    console.log(`--------------------------//------ ${dados.length} Registros ------//--------------------------`)
    

    for (let i = 0; i < dados.length; i++){

        console.log()
        console.log(`\n--------------------------//------Linha ${i}------//--------------------------`);

        const AIT = dados[i]['AIT Process'];
        const ModoEnvio = dados[i]['Modo de Envio'];
        const CarimboLumma = dados[i]['Carimbo Lumma'];
        const CarimboCliente = dados[i]['Carimbo Cliente'];
        const NIC = dados[i]['NIC'];
        const Orgao = dados[i]['Orgao'];
        const Valor = dados[i]['Valor integral'];
        const ValorDesc = dados[i]['Valor com desconto'];
        const Locatario = dados[i]['Loc'];
        const Banco = dados[i]['Banco'];
        const LummaP = dados[i]['Lumma paga']
        const Autorizou = dados[i]['Quem autorizou']
        const Contrato = dados[i]['Contrato']

        // Clica em Validar dados
        await new Promise(resolve => setTimeout(resolve, 1000));
        await page.mouse.click(68, 366);
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Verifica se o sistema está no local certo para começar o loop

        while (contador_inicio < 11){
            const xinicio = 158;
            const yinicio = 401;
            await page.screenshot({ path: 'images/screenshot_inicio.png' });
            const img_inicio = await Jimp.read('images/screenshot_inicio.png');
            const corPixel = Jimp.intToRGBA(img_inicio.getPixelColor(xinicio, yinicio));
            let cor = corPixel.r === 70 && corPixel.g === 83 && corPixel.b === 123;
            if (cor) {
                console.log(`${new Date().toLocaleString()} - Indo para AIT: `, AIT);
                await new Promise(resolve => setTimeout(resolve, 3000));
                contador_inicio = 11;
            } else {
                await new Promise(resolve => setTimeout(resolve, 2000));
                contador_inicio = contador_inicio + 1;
    
                if (contador_inicio > 10) {
                    console.log(`${new Date().toLocaleString()} - Não foi possível carregar o sistema`);
                    console.log('--------------------------//--------------------------//--------------------------');
                    await browser.close();
                }
        }}
        
        await new Promise(resolve => setTimeout(resolve, 5000));

        // AIT 
        await page.mouse.click(295, 345, { clickCount: 2 });
        // Pega dados da coluna no excel 
        await page.keyboard.type(AIT.toString());  
        await page.keyboard.press('Tab');
        await page.keyboard.press('Enter');
        await new Promise(resolve => setTimeout(resolve, 2000));

        // Verifica se a AIT inserida foi encontrada
        let flagReiniciar = false
        let contador_mais = 0;
        while (contador_mais < 4){
            const xmais = 285;
            const ymais = 589;
            await page.screenshot({ path: 'images/screenshot_+.png' });
            const img_mais = await Jimp.read('images/screenshot_+.png');
            const corPixelmais = Jimp.intToRGBA(img_mais.getPixelColor(xmais, ymais));
            let cor_mais = corPixelmais.r === 255 && corPixelmais.g === 255 && corPixelmais.b === 255;
            if (cor_mais) {
                console.log(`${new Date().toLocaleString()} - Encontrou AIT: `, AIT);
                contador_mais = 4;
            } else {
                await new Promise(resolve => setTimeout(resolve, 2000));
                contador_mais = contador_mais + 1;
                if (contador_mais > 3) {
                    console.log(`${new Date().toLocaleString()} - Ait não encontrado: `, AIT);
                    console.log('--------------------------//--------------------------//--------------------------');
                    dados[i]['STATUS'] = 'AIT não encontrada';
                    // Converta o array de objetos atualizado de volta para um worksheet
                    const ws = XLSX.utils.json_to_sheet(dados);
                    // Substitua o worksheet antigo pelo novo no workbook
                    workbook.Sheets[sheet_name_list[0]] = ws;
                    // Escreva o workbook atualizado de volta para o arquivo
                    XLSX.writeFile(workbook, "S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Validacao_retorno.xlsx"); 
                    flagReiniciar = true;
                }
        }}
        if(flagReiniciar) continue;

        // Click +
        await page.mouse.click(260, 578);

        // Verifica se o formulário abriu 
        let flagReload = false
        let contador_abriu = 0;
        while (contador_abriu < 20){
            const xabriu = 140;
            const yabriu = 943;
            await page.screenshot({ path: 'images/screenshot_abriu.png' });
            const img_abriu = await Jimp.read('images/screenshot_abriu.png');
            const corPixel_abriu = Jimp.intToRGBA(img_abriu.getPixelColor(xabriu, yabriu));
            let cor_abriu = corPixel_abriu.r === 255 && corPixel_abriu.g === 255 && corPixel_abriu.b === 255;
            if (!cor_abriu) {
                console.log(`${new Date().toLocaleString()} - Entrou na tela de validação`);
                await new Promise(resolve => setTimeout(resolve, 3500));
                contador_abriu = 20;
            } else {
                await new Promise(resolve => setTimeout(resolve, 2000));
                contador_abriu = contador_abriu + 1;
                if (contador_abriu > 19) {
                    console.log(`${new Date().toLocaleString()} - Não foi possível abrir o formulário`);
                    console.log('--------------------------//--------------------------//--------------------------');
                    await page.reload();
                    flagReload = true
                }
    }}
        if (flagReload) continue;

        // Leitura do Iframe
        const frameElement = await page.$('#frameDetails');
        const frame = await frameElement.contentFrame();


        // Data atual
        await new Promise(resolve => setTimeout(resolve, 500));
        await frame.click('#radioDataAtual')
        await new Promise(resolve => setTimeout(resolve, 500));

    
        // Arruma o array com os locatários
        let companiesArray = process.env.strings.split("\n").map(s => s.trim());
        // Verifica se o elemento está presente
        const isSelectPresent = await frame.$('#cmbLocatarioNot') !== null;
        if (!isSelectPresent) {
            console.log(`${new Date().toLocaleString()} - Um elemento select não está presente na página. Fechando`);
            await browser.close();
            return;
        } 
        // Obtém o valor selecionado no elemento
        const selectedLocatario = await frame.evaluate(() => {
            const selector = document.querySelector('#cmbLocatarioNot');
            return selector.options[selector.selectedIndex].text;
        });
        // Verifica se o locatário selecionado está dentro do array
        let isPresent = companiesArray.includes(selectedLocatario.trim());
        if (isPresent) {
            console.log(`${new Date().toLocaleString()} - Erro: Locatário encontrado na lista.`)
            dados[i]['STATUS'] = 'Erro: Locatário encontrado na lista';
            const ws = XLSX.utils.json_to_sheet(dados);
            workbook.Sheets[sheet_name_list[0]] = ws;
            XLSX.writeFile(workbook, 'S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Validacao_retorno.xlsx');
            await page.reload();
            await new Promise(resolve => setTimeout(resolve, 500));
            continue;
        } 
        await new Promise(resolve => setTimeout(resolve, 500));


        // Checa se LummaPaga está marcado, e arruma caso esteja ou não.
        async function isCheckboxChecked(page, selector) {
            return frame.$eval(selector, el => el.checked);
        }
        const checkboxSelector = '#chkLummaPaga'; 
        const isChecked = await isCheckboxChecked(frame, checkboxSelector);
        if (LummaP.toUpperCase() === 'SIM') {
             if (!isChecked){
                await frame.click('#lblLummaPaga')
             } 
        } else{
            if (isChecked){
                await frame.click('#lblLummaPaga')
            }
        }

        // Quem autoriza
        await frame.click('#txtAutorizacaoPagamento');
        await page.keyboard.type(Autorizou.toString()); 
        await new Promise(resolve => setTimeout(resolve, 2000));

        // NIC
        await frame.click('#cmbTipoNIC');
        await page.keyboard.type(NIC.toString()); 
        await page.keyboard.press('Enter');

        // Tipo contrato
        await frame.click('#cmbTipoContrato');
        await new Promise(resolve => setTimeout(resolve, 500));
        if (Contrato !== undefined) {
            await page.keyboard.type(Contrato.toString());
            await new Promise(resolve => setTimeout(resolve, 1000));
            await page.keyboard.press('Enter');
        } 
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Modo de envio
        await new Promise(resolve => setTimeout(resolve, 2000));
        await frame.click('.campoModoEnvio.inputLong ')
        await page.keyboard.type(ModoEnvio.toString()); 
        await page.keyboard.press('Enter');


        // Carimbo da lumma
        await new Promise(resolve => setTimeout(resolve, 500));
        await frame.click('.DriverNameClass.datepicker');
        await page.keyboard.type(CarimboLumma.toString(), {delay: 100}); 
        await page.keyboard.press('Tab');

        // Carimbo do Cliente
        await new Promise(resolve => setTimeout(resolve, 500));
        await page.keyboard.type(CarimboCliente.toString()); 
        await page.keyboard.press('Enter');
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Locatario 1
        await frame.click('#cmbLocatarioNot');
        if (Locatario !== undefined) {
            await page.keyboard.type(Locatario.toString());
            await new Promise(resolve => setTimeout(resolve, 1000));
            await page.keyboard.press('Enter');
        }
        await new Promise(resolve => setTimeout(resolve, 250));
        await page.mouse.click(594, 391)

        // Locatario 2
        if (Locatario !== undefined) {
            await frame.click('.campoLocatarioBoleto.inputLong.grupoAtual');
            await page.keyboard.type(Locatario.toString());
            await page.keyboard.press('Enter');
        } 
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Órgão

        if (Orgao !== undefined) {
        await new Promise(resolve => setTimeout(resolve, 500));
        await frame.click('#cmbOrgaoBol_chosen');
        await page.keyboard.type(Orgao.toString()); 
        await new Promise(resolve => setTimeout(resolve, 500));
        await page.keyboard.press('Enter');
        }

        // Valor
        await new Promise(resolve => setTimeout(resolve, 500));
        await frame.click('#txtValorBol');
        await new Promise(resolve => setTimeout(resolve, 500));
        await page.keyboard.press('Tab');
        await new Promise(resolve => setTimeout(resolve, 500.));
        await page.keyboard.type(Valor.toString()); 
        await new Promise(resolve => setTimeout(resolve, 500));
        await page.keyboard.press('Tab');


        // Valor comd desconto
        await new Promise(resolve => setTimeout(resolve, 500));
        await page.keyboard.type(ValorDesc.toString()); 
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Banco
        if (Banco !== undefined){
        await frame.evaluate((textToSelect) => {
            let selectElement = document.querySelector("#cmbBancoBol");
            for(let i = 0; i < selectElement.options.length; i++) {
              if(selectElement.options[i].text === textToSelect) {
                selectElement.selectedIndex = i;
                break;
              }
            }
          }, Banco);      
        }    
        await new Promise(resolve => setTimeout(resolve, 2000));
    // Limpa campos de email

        await frame.click('#txtEmailGestorBol')
        await page.keyboard.down('Control');
        await page.keyboard.press('KeyA');
        await page.keyboard.up('Control');
        await page.keyboard.press('Backspace')

        await new Promise(resolve => setTimeout(resolve, 500));
        for (let j=0; j<8; j++) {
          await page.keyboard.press('Tab')
          await page.keyboard.down('Control');
          await page.keyboard.press('KeyA');
          await page.keyboard.up('Control');
          await page.keyboard.press('Backspace')
          await new Promise(resolve => setTimeout(resolve, 500));
        }
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Salvar 1
        await frame.click('.btn.btn-primary ');
        await new Promise(resolve => setTimeout(resolve, 2000));

         // Verificação de erros ao salvar
        let flagRestart = false;
        const xfinal = 1026;
        const yfinal = 403;
        await page.screenshot({ path: `images/screenshot_salvou.png` });
        await new Promise(resolve => setTimeout(resolve, 2000));
        const img_final = await Jimp.read('images/screenshot_salvou.png');
        const corPixel_final = Jimp.intToRGBA(img_final.getPixelColor(xfinal,yfinal));
        let cor_final = corPixel_final.r === 46 && corPixel_final.g === 46 && corPixel_final.b === 46;
        await new Promise(resolve => setTimeout(resolve, 2000))
             if (cor_final) {
                let erro = "S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Erro_AIT_ " + AIT + ".png";
                await page.screenshot({ path: erro})
                await new Promise(resolve => setTimeout(resolve, 1000));
                await page.keyboard.press("Tab");
                await new Promise(resolve => setTimeout(resolve, 1000));
                await page.keyboard.press("Enter");
                await page.mouse.click(199,940);
                console.log(`${new Date().toLocaleString()} - Erro ou falta de informações: `, AIT);
                console.log('--------------------------//--------------------------//--------------------------');
                dados[i]['STATUS'] = 'Erro ou falta de informações';
                const ws = XLSX.utils.json_to_sheet(dados);
                workbook.Sheets[sheet_name_list[0]] = ws;
                XLSX.writeFile(workbook, 'S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Validacao_retorno.xlsx');
                await page.reload();
                flagRestart = true;
             } 

        if(flagRestart) continue;

        // Salvar 2
        await page.keyboard.press('Tab');
        await new Promise(resolve => setTimeout(resolve, 250));
        await page.keyboard.press('Tab');
        await new Promise(resolve => setTimeout(resolve, 250));
        await page.keyboard.press('Tab');
        await new Promise(resolve => setTimeout(resolve, 250));
        await page.keyboard.press('Tab');
        await new Promise(resolve => setTimeout(resolve, 250));
        await page.keyboard.press('Enter');
        

        // Verifica se salvou ou ainda está carregando
        let flagReiniciar2 = false
        let contador_salvou = 0
        while(contador_salvou < 4) {
            const xsalvou = 957;
            const ysalvou = 325;
            await page.screenshot({ path: 'images/screenshot_salvou.png' });
            await new Promise(resolve => setTimeout(resolve, 5000));
            const img_salvou = await Jimp.read('images/screenshot_salvou.png');
            const corPixel_salvou = Jimp.intToRGBA(img_salvou.getPixelColor(xsalvou,ysalvou));
            let cor_salvou = corPixel_salvou.r === 245 && corPixel_salvou.g === 245 && corPixel_salvou.b === 245;
            if (cor_salvou){  
                await page.keyboard.press('Tab');
                await new Promise(resolve => setTimeout(resolve, 250));
                await page.keyboard.press('Enter');
                    contador_salvou = 4;
            } else{
                await new Promise(resolve => setTimeout(resolve, 2000));
                contador_salvou++;

                if (contador_salvou > 3){
                    let erro = "S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Erro_AIT_ " + AIT + ".png";
                    await page.screenshot({ path: erro})
                    await page.keyboard.press('Tab')
                    await new Promise(resolve => setTimeout(resolve, 500));
                    await page.keyboard.press('Enter')
                    await new Promise(resolve => setTimeout(resolve, 500));
                    await page.mouse.click(199,940);
                    await new Promise(resolve => setTimeout(resolve, 500));
                    console.log(`${new Date().toLocaleString()} - Não salvou por erro do sistema, validar manualmente AIT:`, AIT)
                    console.log('--------------------------//--------------------------//--------------------------');
                    dados[i]['STATUS'] = 'Erro no Process';
                    flagReiniciar2 = true;
                    const ws = XLSX.utils.json_to_sheet(dados);
                    workbook.Sheets[sheet_name_list[0]] = ws;
                    XLSX.writeFile(workbook, "S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Validacao_retorno.xlsx"); 
    

                }
            }}

        if (flagReiniciar2) continue;
        console.log(`${new Date().toLocaleString()} - Validação concluída, AIT: `, AIT)
        console.log('--------------------------//--------------------------//--------------------------');
        // No final de cada iteração do loop
        dados[i]['STATUS'] = 'OK'; // Assume que 'dados' é o array de objetos que você leu do Excel
        // Converta o array de objetos atualizado de volta para um worksheet
         const ws = XLSX.utils.json_to_sheet(dados);
         // Substitua o worksheet antigo pelo novo no workbook
         workbook.Sheets[sheet_name_list[0]] = ws;
         // Escreva o workbook atualizado de volta para o arquivo
         XLSX.writeFile(workbook, "S:\\Automacoes\\Multas\\07 - Multas ALD\\Retorno\\Validacao_retorno.xlsx"); 
    
        // Tirar um screenshot e salvar na mesma pasta do script
        await new Promise(resolve => setTimeout(resolve, 2000));
        await page.screenshot({ path: 'images/screenshot_recomeço.png' });
    }
    console.log('--------------------------//--------------------------//--------------------------');
    console.log('--------------------------//--------------------------//--------------------------');
    console.log(`--------------------------//     END AUTOMATION - USER:  ${email}    //------`);
    console.log('--------------------------//--------------------------//--------------------------');
    console.log('--------------------------//--------------------------//--------------------------');

    
    // Fechar o navegador 
    await browser.close();
    // Enviar email ao usuário 
    sendEmail(email, "07 - Multas ALD") 
    console.error("Processo finalizado com sucesso.")
})(); 