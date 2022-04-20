const { chromium } = require('playwright');
const excel = require('exceljs');
let workbook = new excel.Workbook();
let worksheet = workbook.addWorksheet('Sheet1');


worksheet.columns = [
    { header: "#", key: "id2", width: 8 },
    { header: "ชื่อ-สกุล", key: "name2", width: 8 },
    { header: "เพศ", key: "gender2", width: 8 },
    { header: "ประเภทการแข่งขัน", key: "type2", width: 8 },
    { header: "ขนาดเสื้อ", key: "shirtSize2", width: 8 },
    { header: "แจ้งโอน", key: "paidStatus2", width: 8 },
    { header: "วันที่สมัคร", key: "date2", width: 8 },
];

(async () => {

	const browser = await chromium.launch({
		headless: false // false if you can see the browser
	})
	const page = await browser.newPage()

	// navigate and wait until network is idle
	await page.goto('https://www.rotarymagkangmarathon.com/run/userlist.php?e=1&page=1', { waitUntil: 'networkidle' })

	await page.waitForSelector('.page-link') // wait for the element
	// get the elements in pagination
	const numberPages = await page.$$eval('.page-link', numberpages => {
		return numberpages.map((numberPage) => {
			return parseInt(numberPage.innerText)
		})
	})
	// get total pages in pagination
	const totalPages = Math.max(...numberPages.filter((p) => !isNaN(p)))
    console.log(totalPages)
	// get the articles per page
	for (let i = 1; i <= totalPages; i++) {

		try {
            await page.waitForSelector('.body-datatable');
            const articlesPerPage = await page.$$eval('tbody tr', (users) => {
                return users.map(user => {
                    const id = user.querySelector('td:nth-child(1)');
                    const name = user.querySelector('td:nth-child(2)');
                    const gender = user.querySelector('td:nth-child(3)');
                    const type = user.querySelector('td:nth-child(4)');
                    const shirtSize = user.querySelector('td:nth-child(5)');
                    const paidStatus =  user.querySelector('td:nth-child(6)');
                    const date =  user.querySelector('td:nth-child(7)');
                    return {
                        id2: id.textContent.trim(),
                        name2: name.textContent.trim(),
                        gender2: gender.textContent.trim(),
                        type2: type.textContent.trim(),
                        shirtSize2: shirtSize.textContent.trim(),
                        paidStatus2: paidStatus.textContent.trim(),
                        date2: date.textContent.trim()
                    };
                });
            });

            articlesPerPage.forEach(row => {
                worksheet.addRow(row)
            })

			if (i != totalPages) {
				// await page.click(`text='${i+1}'`);

				// for this website, another option to navigate is to use URL 
                await page.goto(`https://www.rotarymagkangmarathon.com/run/userlist.php?e=1&page=${i+1}`, { waitUntil: 'networkidle' })
            }
			
		} catch (error) {
			console.log({ error })
		}

		// optional, to see more clearly how the browser works 
		// wait 4000ms 
		await delay(4000);

	}

	// close page and browser
	await page.close()
	await browser.close()
    await workbook.xlsx.writeFile('data.xlsx')

})();

// function to wait a while
function delay(time) {
	return new Promise(function(resolve) { 
		setTimeout(resolve, time)
	});
 }