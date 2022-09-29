import { chromium } from 'playwright'
import * as XLSX from 'xlsx'
import ProgressBar from 'progress'
import { Command } from 'commander'
import chalk from 'chalk'

const program = new Command()

program
  .option('-u, --user', 'username')
  .option('-p, --password', 'password')
  .option('-l, --link', 'your login page link')
  .option('-d, --debug', 'open debug mode');

program.parse(process.argv);

const options = program.opts();
const [username, password, path] = program.args;

const log = console.log;
const recordAllList = [];
const allFailedList = [];
const parallelCount = 5;
const debug = options.debug || false;

(async () => {
  const browser = await chromium.launch({ headless: !debug });
  const loginContext = await browser.newContext();
  const page = await loginContext.newPage();

  await page.goto(
    "https://login.sina.com.cn/signup/signin.php?entry=general&r=http%3A%2F%2Ftousu.sina.com.cn%2Findex.php%2Fuser%2Fmessage"
  );
  await page.locator('input[name="username"]').fill(username);
  await page.locator('input[name="password"]').fill(password);
  await page.locator('input[type="submit"]').click();

  page.on("response", async (response) => {
    if (response.url().indexOf("/message_list") > -1) {
      const text = await response.text();
      const total = debug ? 10 : Number(text.split('"item_count":')[1].split("}}")[0]);
      const listUrl = response
        .url()
        .replace("page_size=10", `page_size=${total}`);
      const responseList = await loginContext.request.get(listUrl);

      const bar = new ProgressBar(" 正在处理： [:bar] :current/:total", {
        complete: "█",
        incomplete: "░",
        width: 40,
        total,
      });
      const newText = await responseList.text();
      const details = newText
        .split("https:")
        .map((item) => item.substring(4, 51).replaceAll("\\", ""))
        .slice(1);

      // 全部标记为已读
      if (!debug) {
        const pageForDetail = await loginContext.newPage();
        await pageForDetail.goto(
          "https://tousu.sina.com.cn/index.php/user/message"
        );
        const read = await pageForDetail.locator(".readBtn").filter({
          hasText: "全部已读",
        });
        await read.click();
      }

      // 并行处理
      const tasks = Math.floor(details.length / parallelCount);
      const promises = [];
      for (let i = 0; i <= parallelCount; i++) {
        const curPage = await loginContext.newPage();
        const list = details.slice(i * tasks, (i + 1) * tasks);
        promises.push(extractor(list, curPage, bar));
      }

      await Promise.all(promises);

      // 失败重试
      if (allFailedList.length) {
        const curPage = await loginContext.newPage();
        await extractor(allFailedList, curPage, bar, true);
      }

      const ws = XLSX.utils.json_to_sheet(recordAllList);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      XLSX.utils.sheet_add_aoa(
        ws,
        [
          [
            "投诉编号",
            "投诉对象",
            "投诉问题",
            "投诉要求",
            "涉诉金额",
            "联系方式",
            "投诉进度",
            "链接",
            "商家处理时间",
            "申请完成时间1",
            "申请完成时间2",
            "投诉内容",
          ],
        ],
        { origin: "A1" }
      );

      /* calculate column width */
      // const max_width = recordAllList.reduce((w, r) => Math.max(w, r['1'].length), 20);
      ws["!cols"] = [{ wch: 20 }];
      const [, root, user] = process.cwd().split('/')
      const exportPath = `/${root}/${user}/Downloads`
      XLSX.writeFile(wb, path || exportPath + "/Report.xlsx");

      await browser.close();
    }
  });
})();

async function extractor(detailList, page, bar, retry = false) {
  let currentRecord;
  for (let view of detailList) {
    try {
      await page.goto(`https://${view}`, {
        timeout: 50 * 1000,
      });

      const ulList = await page.locator(".ts-q-list").textContent();
      const record = ulList.replace(/\s{1,}/g, " ");
      currentRecord = record.split("：").reduce((accu, current, index) => {
        if (index === 0) {
          return accu;
        }
        accu[index] = current.split(" ")[1] || current.split(" ")[0];
        return accu;
      }, {});

      // 获取黑猫处理时间
      const blackCatDate = await page
        .locator(".ts-d-user", {
          has: page.locator("text=黑猫消费者服务平台"),
        })
        .first()
        .locator("span >> nth=-1")
        .textContent();
      currentRecord["8"] = view;
      currentRecord["9"] = blackCatDate;

      // 获取商家处理时间
      const selector = await page.locator(".ts-d-user", {
        has: page.locator("text=申请完成投诉"),
      });
      const count = await selector.count();
      let firstDate = "";
      let secondDate = "";
      if (count !== 0) {
        firstDate = await selector
          .last()
          .locator("span >> nth=-1")
          .textContent();
      }
      if (count === 2) {
        secondDate = await selector
          .first()
          .locator("span >> nth=-1")
          .textContent();
      }
      currentRecord["10"] = firstDate;
      currentRecord["11"] = secondDate;

      // 获取投诉内容
      const isTeamComplain =
        (await page
          .locator(".ts-d-user")
          .filter({
            hasText: "发起集体投诉",
          })
          .count()) > 0;

      const filterText = isTeamComplain ? "参与集体投诉" : "发起投诉";
      let complainSelector = await page
        .locator(".ts-d-item")
        .filter({
          hasText: filterText,
        })
        .last();
      if (!complainSelector) {
        throw new Error("can not get the element");
      }
      const content = await complainSelector
        .locator(".ts-reply")
        .locator("p >> nth=1")
        .textContent();
      currentRecord["12"] = content;
      recordAllList.push(currentRecord);
      bar.tick(1);
    } catch (err) {
      try {
        const complainSelector = await page
          .locator(".ts-d-item")
          .filter({
            hasText: "发起投诉",
          })
          .last();
        const content = await complainSelector
          .locator(".ts-reply")
          .locator("p >> nth=1")
          .textContent();
        currentRecord["12"] = content;
        recordAllList.push(currentRecord);
        bar.tick(1);
      } catch (err) {
        allFailedList.push(view);
        if (retry) {
          log(
            chalk.red(
              "\n提取失败, 请手动处理: ",
              chalk.white(`https://${view}`)
            )
          );
        }
      }
    }
  }
}
