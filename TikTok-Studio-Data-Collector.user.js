// ==UserScript==
// @name         TikTok Studio 数据采集器
// @namespace    qualitell.tiktok.collector
// @version      0.2.0
// @description  批量采集 TikTok Studio 公开视频的播放、观看时长、完播率、留存和流量来源，并导出 CSV；自动排除预约及非公开视频。
// @author       Qualitell
// @homepageURL  https://github.com/piboss001/tiktok-studio-data-collector
// @supportURL   https://github.com/piboss001/tiktok-studio-data-collector/issues
// @match        https://www.tiktok.com/tiktokstudio
// @match        https://www.tiktok.com/tiktokstudio/*
// @run-at       document-start
// @grant        unsafeWindow
// @updateURL    https://raw.githubusercontent.com/piboss001/tiktok-studio-data-collector/main/TikTok-Studio-Data-Collector.user.js
// @downloadURL  https://raw.githubusercontent.com/piboss001/tiktok-studio-data-collector/main/TikTok-Studio-Data-Collector.user.js
// ==/UserScript==

(() => {
  'use strict';

  const page = typeof unsafeWindow !== 'undefined'
    ? unsafeWindow
    : window;

  const VERSION = '0.2.0';

  const items = new Map();
  const detailRows = new Map();
  const privateIds = new Set();
  const unknownIds = new Set();

  let failedCount = 0;
  let busy = false;

  const originalFetch = page.fetch.bind(page);
  const OriginalXHR = page.XMLHttpRequest;

  const sleep = ms =>
    new Promise(resolve => setTimeout(resolve, ms));

  const $ = id =>
    document.getElementById(id);

  function num(value) {
    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {
      return '';
    }

    const n = Number(value);

    return Number.isNaN(n)
      ? ''
      : n;
  }

  function formatPercent(value) {
    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {
      return '';
    }

    const n = Number(value);

    if (Number.isNaN(n)) {
      return '';
    }

    return (n * 100).toFixed(1) + '%';
  }

  function formatDate(timestamp) {
    if (!timestamp) {
      return '';
    }

    return new Date(
      Number(timestamp) * 1000
    ).toLocaleString('zh-CN');
  }

  function csvEscape(value) {
    if (
      value === null ||
      value === undefined
    ) {
      return '';
    }

    const str = String(value);

    if (/[",\n\r]/.test(str)) {
      return '"' +
        str.replace(/"/g, '""') +
        '"';
    }

    return str;
  }

  function safeClone(value) {
    try {
      return JSON.parse(
        JSON.stringify(value)
      );
    } catch {
      return null;
    }
  }

  function isScheduled(item) {
    const now =
      Math.floor(Date.now() / 1000);

    const postTime =
      Number(item?.post_time || 0);

    const scheduleTime =
      Number(item?.schedule_time || 0);

    return (
      postTime > now + 60 ||
      scheduleTime > now + 60
    );
  }

  function candidateItems() {
    return [...items.values()]
      .filter(item => !isScheduled(item));
  }

  function scheduledItems() {
    return [...items.values()]
      .filter(item => isScheduled(item));
  }

  /* ==============================
     捕获 TikTok 作品列表
  ============================== */

  function captureItemList(url, data) {
    if (
      !url ||
      !String(url).includes(
        '/tiktok/creator/manage/item_list/v1/'
      )
    ) {
      return;
    }

    if (
      !data ||
      !Array.isArray(data.item_list)
    ) {
      return;
    }

    for (const item of data.item_list) {
      if (!item?.item_id) {
        continue;
      }

      items.set(
        String(item.item_id),
        safeClone(item)
      );
    }

    if (data.has_more) {
      updateUI(
        `已发现 ${items.size} 条，继续加载…`
      );
    } else {
      updateUI(
        `已发现 ${items.size} 条`
      );
    }
  }

  /* ==============================
     Hook fetch
  ============================== */

  page.fetch = async function (...args) {
    const response =
      await originalFetch(...args);

    try {
      const input = args[0];

      const url =
        typeof input === 'string'
          ? input
          : input?.url;

      if (
        url &&
        String(url).includes(
          '/tiktok/creator/manage/item_list/v1/'
        )
      ) {
        response
          .clone()
          .json()
          .then(data => {
            captureItemList(
              url,
              data
            );
          })
          .catch(() => {});
      }
    } catch {}

    return response;
  };

  /* ==============================
     Hook XMLHttpRequest
  ============================== */

  if (OriginalXHR?.prototype) {
    const originalOpen =
      OriginalXHR.prototype.open;

    const originalSend =
      OriginalXHR.prototype.send;

    OriginalXHR.prototype.open =
      function (
        method,
        url,
        ...rest
      ) {
        this.__qualitellTikTokURL =
          url;

        return originalOpen.call(
          this,
          method,
          url,
          ...rest
        );
      };

    OriginalXHR.prototype.send =
      function (...args) {
        this.addEventListener(
          'load',
          () => {
            try {
              const url =
                this.__qualitellTikTokURL;

              if (
                url &&
                String(url).includes(
                  '/tiktok/creator/manage/item_list/v1/'
                )
              ) {
                const data =
                  JSON.parse(
                    this.responseText
                  );

                captureItemList(
                  url,
                  data
                );
              }
            } catch {}
          }
        );

        return originalSend.apply(
          this,
          args
        );
      };
  }

  /* ==============================
     CSRF
  ============================== */

  function getCsrfToken() {
    const entry =
      document.cookie
        .split('; ')
        .find(cookie =>
          cookie.startsWith(
            'tt_csrf_token='
          )
        );

    if (!entry) {
      return '';
    }

    const value =
      entry
        .split('=')
        .slice(1)
        .join('=');

    try {
      return decodeURIComponent(
        value
      );
    } catch {
      return value;
    }
  }

  /* ==============================
     Insight API
  ============================== */

  async function fetchInsight(videoId) {
    const typeRequests = [
      {
        insigh_type: 'video_info',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_traffic_source_percent_realtime',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_retention_rate_realtime',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_view_realtime',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_total_duration_realtime',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_per_duration_realtime',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_finish_rate_realtime',
        aweme_id: videoId
      },
      {
        insigh_type:
          'video_new_follower_realtime',
        aweme_id: videoId
      }
    ];

    const lang =
      document.documentElement.lang ||
      'zh-Hans';

    const timezone =
      Intl
        .DateTimeFormat()
        .resolvedOptions()
        .timeZone ||
      'Asia/Shanghai';

    const tzOffset =
      -new Date()
        .getTimezoneOffset() *
      60;

    const params =
      new URLSearchParams({
        locale: lang,
        aid: '1988',
        priority_region: 'US',
        region: 'US',
        tz_name: timezone,
        app_name:
          'tiktok_creator_center',
        app_language: lang,
        device_platform:
          'web_pc',
        channel:
          'tiktok_web',
        tz_offset:
          String(tzOffset),
        type_requests:
          JSON.stringify(
            typeRequests
          )
      });

    const csrf =
      getCsrfToken();

    const headers = {
      accept:
        'application/json, text/plain, */*'
    };

    if (csrf) {
      headers['tt-csrf-token'] =
        csrf;
    }

    const response =
      await originalFetch(
        '/aweme/v2/data/insight/?' +
          params.toString(),
        {
          credentials:
            'include',
          headers
        }
      );

    const data =
      await response.json();

    return {
      httpStatus:
        response.status,
      data
    };
  }

  /* ==============================
     留存
  ============================== */

  function retentionAt(
    list,
    milliseconds
  ) {
    const hit =
      (list || []).find(
        item =>
          String(
            item.timestamp
          ) ===
          String(
            milliseconds
          )
      );

    if (!hit) {
      return '';
    }

    return Number(
      hit.value
    );
  }

  /* ==============================
     流量来源
  ============================== */

  function trafficToObject(list) {
    const result = {};

    for (const item of list || []) {
      result[item.key] =
        Number(item.value);
    }

    return result;
  }

  /* ==============================
     整理单条数据
  ============================== */

  function normalizeRow(
    item,
    insight
  ) {
    const data =
      insight.data || {};

    const videoInfo =
      data.video_info || {};

    const statistics =
      videoInfo.statistics || {};

    const retention =
      data
        .video_retention_rate_realtime
        ?.value
        ?.list || [];

    const traffic =
      trafficToObject(
        data
          .video_traffic_source_percent_realtime
          ?.value
          ?.value || []
      );

    const durationMs =
      Number(
        videoInfo.video?.duration ||
        item.duration ||
        0
      );

    const durationSec =
      durationMs
        ? durationMs / 1000
        : '';

    const averageWatch =
      num(
        data
          .video_per_duration_realtime
          ?.value
          ?.value
      );

    const views =
      num(
        data
          .realtime_total_video_views
          ?.value
          ?.value ??
        item.play_count ??
        statistics.play_count
      );

    const likes =
      num(
        statistics.digg_count ??
        item.like_count
      );

    const comments =
      num(
        statistics.comment_count ??
        item.comment_count
      );

    const shares =
      num(
        statistics.share_count ??
        item.share_count
      );

    const favorites =
      num(
        statistics.collect_count ??
        item.favorite_count
      );

    return {
      account:
        videoInfo.author
          ?.unique_id ||
        '',

      nickname:
        videoInfo.author
          ?.nickname ||
        '',

      video_id:
        String(
          item.item_id ||
          videoInfo.aweme_id ||
          ''
        ),

      title:
        item.desc ||
        videoInfo.desc ||
        '',

      publish_time:
        formatDate(
          item.post_time ||
          videoInfo.create_time
        ),

      duration_sec:
        durationSec === ''
          ? ''
          : Number(
              durationSec.toFixed(
                3
              )
            ),

      views,

      likes,

      comments,

      shares,

      favorites,

      new_followers:
        num(
          data
            .realtime_new_followers
            ?.value
            ?.value
        ),

      avg_watch_sec:
        averageWatch,

      watch_ratio:
        (
          averageWatch !== '' &&
          durationSec
        )
          ? averageWatch /
            durationSec
          : '',

      finish_rate:
        num(
          data
            .video_finish_rate_realtime
            ?.value
            ?.value
        ),

      total_watch_sec:
        num(
          data
            .video_total_duration_realtime
            ?.value
            ?.value
        ),

      retention_1s:
        retentionAt(
          retention,
          1000
        ),

      retention_2s:
        retentionAt(
          retention,
          2000
        ),

      retention_3s:
        retentionAt(
          retention,
          3000
        ),

      retention_5s:
        retentionAt(
          retention,
          5000
        ),

      retention_10s:
        retentionAt(
          retention,
          10000
        ),

      for_you:
        traffic['For You'] ??
        '',

      personal_profile:
        traffic[
          'Personal Profile'
        ] ?? '',

      search:
        traffic.Search ??
        '',

      follow:
        traffic.Follow ??
        '',

      direct_message:
        traffic[
          'Direct Message'
        ] ?? '',

      sound:
        traffic.Sound ??
        '',

      others:
        traffic.Others ??
        '',

      like_rate:
        views
          ? likes / views
          : '',

      engagement_rate:
        views
          ? (
              likes +
              comments +
              shares +
              favorites
            ) / views
          : ''
    };
  }

  /* ==============================
     UI CSS
  ============================== */

  function injectStyle() {
    if (
      document.getElementById(
        'qualitell-tiktok-style'
      )
    ) {
      return;
    }

    const style =
      document.createElement(
        'style'
      );

    style.id =
      'qualitell-tiktok-style';

    style.textContent = `
      #qualitell-tiktok-panel,
      #qualitell-tiktok-panel * {
        box-sizing:border-box;
      }

      #qualitell-tiktok-panel {
        position:fixed;
        right:20px;
        bottom:20px;
        z-index:2147483647;
        width:330px;
        background:rgba(22,22,26,.98);
        color:#fff;
        border:1px solid rgba(255,255,255,.12);
        border-radius:14px;
        overflow:hidden;
        box-shadow:0 15px 45px rgba(0,0,0,.35);
        font-family:Arial,"Microsoft YaHei",sans-serif;
      }

      .qtk-header {
        display:flex;
        justify-content:space-between;
        align-items:center;
        padding:14px 16px;
        border-bottom:1px solid rgba(255,255,255,.08);
      }

      .qtk-title {
        font-size:15px;
        font-weight:700;
      }

      .qtk-version {
        margin-top:3px;
        font-size:11px;
        color:rgba(255,255,255,.55);
      }

      .qtk-minimize {
        border:0;
        background:transparent;
        color:#fff;
        font-size:20px;
        cursor:pointer;
      }

      #qtk-body {
        padding:14px 16px 16px;
      }

      .qtk-row {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:12px;
        padding:5px 0;
        font-size:12px;
      }

      .qtk-row span {
        color:rgba(255,255,255,.62);
      }

      .qtk-row b {
        text-align:right;
        font-weight:600;
      }

      .qtk-buttons {
        display:grid;
        gap:8px;
        margin-top:12px;
      }

      .qtk-buttons button {
        width:100%;
        border:0;
        border-radius:9px;
        padding:10px 12px;
        font-size:13px;
        font-weight:700;
        cursor:pointer;
      }

      #qtk-scan {
        background:#fe2c55;
        color:#fff;
      }

      #qtk-fetch {
        background:#25f4ee;
        color:#111;
      }

      #qtk-export {
        background:#fff;
        color:#111;
      }

      #qtk-reset {
        background:#393940;
        color:#fff;
      }

      .qtk-buttons button:disabled {
        opacity:.35;
        cursor:not-allowed;
      }

      .qtk-progress {
        height:7px;
        background:rgba(255,255,255,.1);
        border-radius:999px;
        margin-top:12px;
        overflow:hidden;
      }

      #qtk-progress-bar {
        height:100%;
        width:0%;
        background:#25f4ee;
        transition:width .2s ease;
      }

      .qtk-tip {
        margin-top:10px;
        font-size:11px;
        line-height:1.55;
        color:rgba(255,255,255,.48);
      }
    `;

    (
      document.head ||
      document.documentElement
    ).appendChild(
      style
    );
  }

  /* ==============================
     UI
  ============================== */

  function buildPanel() {
    if (
      !document.documentElement ||
      document.getElementById(
        'qualitell-tiktok-panel'
      )
    ) {
      return;
    }

    injectStyle();

    const panel =
      document.createElement(
        'div'
      );

    panel.id =
      'qualitell-tiktok-panel';

    panel.innerHTML = `
      <div class="qtk-header">

        <div>
          <div class="qtk-title">
            TikTok 数据采集器
          </div>

          <div class="qtk-version">
            V${VERSION} · 仅统计公开视频
          </div>
        </div>

        <button
          id="qtk-minimize"
          class="qtk-minimize"
        >
          −
        </button>

      </div>

      <div id="qtk-body">

        <div class="qtk-row">
          <span>已发现作品</span>
          <b id="qtk-all">0</b>
        </div>

        <div class="qtk-row">
          <span>预约发布</span>
          <b id="qtk-scheduled">0</b>
        </div>

        <div class="qtk-row">
          <span>待检查已发布</span>
          <b id="qtk-candidates">0</b>
        </div>

        <div class="qtk-row">
          <span>公开已发布</span>
          <b id="qtk-public">0</b>
        </div>

        <div class="qtk-row">
          <span>仅自己/非公开</span>
          <b id="qtk-private">0</b>
        </div>

        <div class="qtk-row">
          <span>失败/未知</span>
          <b id="qtk-failed">0</b>
        </div>

        <div class="qtk-row">
          <span>状态</span>
          <b id="qtk-status">
            等待 TikTok 数据
          </b>
        </div>

        <div class="qtk-buttons">

          <button id="qtk-scan">
            扫描全部作品
          </button>

          <button
            id="qtk-fetch"
            disabled
          >
            获取公开视频数据
          </button>

          <button
            id="qtk-export"
            disabled
          >
            导出 CSV
          </button>

          <button id="qtk-reset">
            清空重扫
          </button>

        </div>

        <div class="qtk-progress">
          <div id="qtk-progress-bar"></div>
        </div>

        <div class="qtk-tip">
          自动排除预约发布、仅自己可见和其它非公开视频。
          首次安装或更新后请刷新 TikTok Studio 作品页。
        </div>

      </div>
    `;

    document
      .documentElement
      .appendChild(panel);

    $('qtk-minimize')
      .addEventListener(
        'click',
        () => {
          const body =
            $('qtk-body');

          body.style.display =
            body.style.display ===
            'none'
              ? ''
              : 'none';
        }
      );

    $('qtk-scan')
      .addEventListener(
        'click',
        scanAll
      );

    $('qtk-fetch')
      .addEventListener(
        'click',
        fetchAllInsights
      );

    $('qtk-export')
      .addEventListener(
        'click',
        exportCsv
      );

    $('qtk-reset')
      .addEventListener(
        'click',
        resetAll
      );

    updateUI('已就绪');
  }

  function updateUI(status) {
    if (
      !$(
        'qualitell-tiktok-panel'
      )
    ) {
      return;
    }

    $('qtk-all').textContent =
      items.size;

    $('qtk-scheduled').textContent =
      scheduledItems().length;

    $('qtk-candidates').textContent =
      candidateItems().length;

    $('qtk-public').textContent =
      detailRows.size;

    $('qtk-private').textContent =
      privateIds.size;

    $('qtk-failed').textContent =
      failedCount +
      unknownIds.size;

    if (status) {
      $('qtk-status').textContent =
        status;
    }

    $('qtk-scan').disabled =
      busy;

    $('qtk-fetch').disabled =
      busy ||
      candidateItems().length === 0;

    $('qtk-export').disabled =
      detailRows.size === 0;
  }

  /* ==============================
     扫描
  ============================== */

  async function scanAll() {
    if (busy) {
      return;
    }

    busy = true;

    updateUI(
      '正在扫描作品…'
    );

    let stableRounds = 0;
    let lastCount =
      items.size;

    let lastHeight =
      document
        .documentElement
        .scrollHeight;

    for (
      let i = 0;
      i < 100;
      i++
    ) {
      page.scrollTo({
        top:
          document
            .documentElement
            .scrollHeight,

        behavior:
          'smooth'
      });

      await sleep(850);

      const newCount =
        items.size;

      const newHeight =
        document
          .documentElement
          .scrollHeight;

      if (
        newCount ===
          lastCount &&
        newHeight ===
          lastHeight
      ) {
        stableRounds++;
      } else {
        stableRounds = 0;
      }

      lastCount =
        newCount;

      lastHeight =
        newHeight;

      updateUI(
        `扫描中：${newCount} 条`
      );

      if (
        stableRounds >= 6
      ) {
        break;
      }
    }

    busy = false;

    if (items.size) {
      updateUI(
        `扫描完成：${items.size} 条`
      );
    } else {
      updateUI(
        '未发现作品，请刷新作品页'
      );
    }
  }

  /* ==============================
     批量获取详细数据
  ============================== */

  async function fetchAllInsights() {
    if (
      busy ||
      !items.size
    ) {
      return;
    }

    busy = true;

    failedCount = 0;

    detailRows.clear();
    privateIds.clear();
    unknownIds.clear();

    const list =
      candidateItems();

    let completed = 0;

    $('qtk-progress-bar')
      .style.width =
      '0%';

    updateUI(
      '开始获取详细数据…'
    );

    for (const item of list) {
      const videoId =
        String(
          item.item_id
        );

      try {
        const insight =
          await fetchInsight(
            videoId
          );

        if (
          insight.httpStatus !==
            200 ||
          insight.data
            ?.status_code !== 0
        ) {
          throw new Error(
            'Insight request failed'
          );
        }

        const privateStatus =
          insight.data
            ?.video_info
            ?.status
            ?.private_status;

        /*
          已验证：
          private_status === 0
          为公开作品。

          其它状态全部视为
          非公开，不导出。
        */

        if (
          privateStatus === 0
        ) {
          detailRows.set(
            videoId,
            normalizeRow(
              item,
              insight
            )
          );

        } else if (
          privateStatus ===
            null ||
          privateStatus ===
            undefined
        ) {
          unknownIds.add(
            videoId
          );

        } else {
          privateIds.add(
            videoId
          );
        }

      } catch (error) {
        console.warn(
          '[TikTok Collector]',
          videoId,
          error
        );

        failedCount++;
      }

      completed++;

      const progress =
        list.length
          ? Math.round(
              completed /
              list.length *
              100
            )
          : 0;

      $('qtk-progress-bar')
        .style.width =
        progress + '%';

      updateUI(
        `详情 ${completed}/${list.length}`
      );

      /*
        不要瞬间大量请求 TikTok。
      */

      await sleep(800);
    }

    busy = false;

    updateUI(
      `完成：公开 ${detailRows.size}，非公开 ${privateIds.size}，失败/未知 ${failedCount + unknownIds.size}`
    );
  }

  /* ==============================
     清空
  ============================== */

  function resetAll() {
    if (busy) {
      return;
    }

    items.clear();
    detailRows.clear();
    privateIds.clear();
    unknownIds.clear();

    failedCount = 0;

    if (
      $('qtk-progress-bar')
    ) {
      $('qtk-progress-bar')
        .style.width =
        '0%';
    }

    updateUI(
      '已清空，请刷新页面重新扫描'
    );
  }

  /* ==============================
     CSV
  ============================== */

  function exportCsv() {
    const rows =
      [...detailRows.values()];

    if (!rows.length) {
      return;
    }

    const columns = [
      ['账号', 'account'],
      ['昵称', 'nickname'],
      ['Video ID', 'video_id'],
      ['标题', 'title'],
      ['发布时间', 'publish_time'],

      ['视频时长(s)', 'duration_sec'],

      ['播放量', 'views'],
      ['点赞', 'likes'],
      ['评论', 'comments'],
      ['分享', 'shares'],
      ['收藏', 'favorites'],

      ['新增粉丝', 'new_followers'],

      ['平均观看时间(s)', 'avg_watch_sec'],
      ['观看倍率', 'watch_ratio'],
      ['完播率', 'finish_rate'],
      ['总观看时间(s)', 'total_watch_sec'],

      ['1秒留存', 'retention_1s'],
      ['2秒留存', 'retention_2s'],
      ['3秒留存', 'retention_3s'],
      ['5秒留存', 'retention_5s'],
      ['10秒留存', 'retention_10s'],

      ['For You', 'for_you'],
      ['个人主页', 'personal_profile'],
      ['搜索', 'search'],
      ['关注', 'follow'],
      ['私信', 'direct_message'],
      ['音乐', 'sound'],
      ['其它', 'others'],

      ['点赞率', 'like_rate'],
      ['互动率', 'engagement_rate']
    ];

    const percentKeys =
      new Set([
        'watch_ratio',
        'finish_rate',

        'retention_1s',
        'retention_2s',
        'retention_3s',
        'retention_5s',
        'retention_10s',

        'for_you',
        'personal_profile',
        'search',
        'follow',
        'direct_message',
        'sound',
        'others',

        'like_rate',
        'engagement_rate'
      ]);

    const lines = [];

    lines.push(
      columns
        .map(
          ([label]) =>
            csvEscape(label)
        )
        .join(',')
    );

    for (const row of rows) {
      lines.push(
        columns
          .map(
            ([, key]) => {
              const value =
                percentKeys.has(
                  key
                )
                  ? formatPercent(
                      row[key]
                    )
                  : row[key];

              return csvEscape(
                value
              );
            }
          )
          .join(',')
      );
    }

    /*
      UTF-8 BOM：
      防止 Excel 打开中文/Emoji乱码。
    */

    const blob =
      new Blob(
        [
          '\uFEFF' +
          lines.join('\r\n')
        ],
        {
          type:
            'text/csv;charset=utf-8'
        }
      );

    const url =
      URL.createObjectURL(
        blob
      );

    const link =
      document.createElement(
        'a'
      );

    const account =
      rows[0]?.account ||
      'tiktok';

    const date =
      new Date()
        .toISOString()
        .slice(0, 10);

    link.href =
      url;

    link.download =
      `${account}_TikTok公开视频数据_${date}.csv`;

    document.body
      .appendChild(
        link
      );

    link.click();

    link.remove();

    setTimeout(
      () => {
        URL.revokeObjectURL(
          url
        );
      },
      3000
    );
  }

  /* ==============================
     启动
  ============================== */

  function boot() {
    if (
      !document.documentElement
    ) {
      setTimeout(
        boot,
        50
      );

      return;
    }

    if (
      document.readyState ===
      'loading'
    ) {
      document.addEventListener(
        'DOMContentLoaded',
        buildPanel,
        {
          once: true
        }
      );
    } else {
      buildPanel();
    }
  }

  boot();

})();
