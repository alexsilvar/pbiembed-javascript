/********************************************/
/*Desenvolvido por Alexsander Ramos da Silva*/
/*	V1.0 DATA:02/02/2018	            */
/*					    */
/*	Nome: Power BI Easy Embed	    */
/*					    */
/********************************************/

//Declaracoes de variaveis
//Objeto que cuida do loguin e requests de acesso
var ADAL;
//Pega o Hash apos o login
var isCallback;
//Usuario logado
var user;

/*Exemplo de uso - Ao carregar o documento, iniciar a carregar o elemento na pagina*/
$(document).ready(function () {
	dados = {
		appId :'16969bc9-APPID-53c458ee7290',		//ID do aplicadivo registrdo no active directory
		hostUrl : 'http://localhost/powerBIEmbed/',			//Endereço cadastrado no active directory e do aplicativo
		groupId : '20b3d911-GROUPID-d25efd102d6f',	//Grupo em que o elemento se encontra (facilmente encontrado na URL)
		tipo : 'dashboard',									//Tipo do elemento a ser exibido
		nome : 'Testando',									//nome do elemento
		container : 'container',							//ID de onde ele cairá no html
	}
    inicializar(dados);   
});



/*
 *	@param appID: ID da aplicacao criada no active directory ou no create app do powerBI
 *	@param adress: caminho como explicitado no redirect link da aplicacao
 *	O caminho do index deve estar explicitado no active directory
 */
function inicializar(dados) {
    ADAL = new AuthenticationContext({
        instance: 'https://login.microsoftonline.com/', //Padrao para login
        tenant: 'common', //Tenant do microsoft azure active directory
        clientId: dados.appId, //ID da aplicacao
        postLogoutRedirectUri: dados.hostUrl //window.location.origin,//Testar se é necessário esse parametro
        //callback: userSignedIn,
        //popUp: true
    });
    isCallback = ADAL.isCallback(window.location.origin/*hash*/);
    ADAL.handleWindowCallback();
    user = ADAL.getCachedUser();
	
	 if (!ADAL.getCachedUser()) {
        conectar();
    }
	//Atualiza os datasets e depois exibe o solicitado
	show(dados);
}

function conectar() {
    ADAL.login();
}

function desconectar() {
    ADAL.logOut();
}
/*
*@param dados contendo informações para correr o fluxo	
*/
function show(dados){
	embedThing(dados.tipo, dados.nome, dados.container);
}

/*
 *	@param thingType: tipo do elemento - 'dashboard' 'report' - escreva corretamente
 *	@param thingName: nome do elemento - dashboard chamado 'Dashboard Interessante'
 * 	@param container: id local onde ficara o elemento - fara um getElementById(container)
 *	So deve ser chamado apos o login ou nao sera possivel conectar ao elemento
 */
function embedThing(thingType, thingName, container) {
    //Busca token de acesso
    ADAL.acquireToken("https://analysis.windows.net/powerbi/api", function (error, token) {
        //Verifica erro de busca no token
        if (error || !token) {
            console.log('ADAL Error Occurred: ' + error);
            return;
        }

        var request = new XMLHttpRequest();
        //Cabecalho da request
        var authHeader = "Bearer " + token;

        request.open('GET', 'https://api.powerbi.com/v1.0/myorg/' + thingType + 's');

        request.setRequestHeader('Authorization', authHeader);

        request.onreadystatechange = function (reposta) {
            if (this.readyState === 4) {
                console.log('Status:', this.status);
                console.log('Body:', this.responseText);
                //responseText contem o formato odata com informacoes importantes
                embing(this.responseText, thingName, thingType, container);
            }
        };
        //envia request com o cabecalho
        request.send();
    });
}


function embing(response, thingName, thingType, container) {
    ADAL.acquireToken("https://analysis.windows.net/powerbi/api", function (error, token) {
        //Transforma o JSON da resposta para formato object e procura o elemento correto pelo nome
        if (!response) {
            console.log('Response:', 'ERRO: Resposta errada');
            return;
        }
        var odata = JSON.parse(response);



        var embed = null;
        for (i = 0; i < odata.value.length; i++) {
            if (odata.value[i].name === thingName || odata.value[i].displayName === thingName) {
                embed = odata.value[i];
                break;
            }
        }
        if (embed === null) {
            console.log('Erro: nome de dashboard|report nao encontrado');
            return;
        }
        //Configuracoes obtidas pela busca
        var models = window['powerbi-client'].models;
        var config = {
            id: embed.id,
            type: thingType,
            accessToken: token,
            embedUrl: embed.embedUrl,
            pageView: "fitToWidth",
            settings: {
                layoutType: models.LayoutType.Custom,
                customLayout: {
                    displayOption: models.DisplayOption.FitToPage
                },
                navContentPaneEnabled: false,
				filterPaneEnabled: false
            }
        };
        // pega o container que conterá o html - tamanho e outras coisas podem ser ditadas via css
        var thingContainer = document.getElementById(container);

        // mostra de fato o elemento dentro do container
        var dashboard = powerbi.embed(thingContainer, config);
        dashboard.iframe.style.border = "none";
        dashboard.iframe.style.width = "100%";
        dashboard.iframe.style.height = "87%";
    });
}
