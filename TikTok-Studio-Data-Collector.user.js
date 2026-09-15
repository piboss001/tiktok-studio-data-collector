// ==UserScript==
// @name         TikTok Studio 数据采集器
// @namespace    qualitell.tiktok.collector
// @version      0.3.2
// @description  批量采集 TikTok Studio 公开视频数据，生成周报、TOP 排名与留存分析；导出美化 XLSX（含视频封面与可点击链接）及 CSV。
// @author       Qualitell
// @homepageURL  https://github.com/piboss001/tiktok-studio-data-collector
// @supportURL   https://github.com/piboss001/tiktok-studio-data-collector/issues
// @match        https://www.tiktok.com/tiktokstudio
// @match        https://www.tiktok.com/tiktokstudio/*
// @run-at       document-start
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @connect      *
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @updateURL    https://raw.githubusercontent.com/piboss001/tiktok-studio-data-collector/main/TikTok-Studio-Data-Collector.user.js
// @downloadURL  https://raw.githubusercontent.com/piboss001/tiktok-studio-data-collector/main/TikTok-Studio-Data-Collector.user.js
// ==/UserScript==

(() => {
  'use strict';

  const page =
    typeof unsafeWindow !== 'undefined'
      ? unsafeWindow
      : window;

  const VERSION = '0.3.2';

  const items = new Map();
  const detailRows = new Map();
  const privateIds = new Set();
  const unknownIds = new Set();
  const coverCache = new Map();

  let failedCount = 0;
  let busy = false;
  let minimized = false;

  const originalFetch =
    page.fetch.bind(page);

  const OriginalXHR =
    page.XMLHttpRequest;

  const sleep = ms =>
    new Promise(resolve =>
      setTimeout(resolve, ms)
    );

  const $ = id =>
    document.getElementById(id);

  const COLORS = {
    navy: 'FF0F172A',
    navy2: 'FF1E293B',
    blue: 'FF2563EB',
    blueLight: 'FFDBEAFE',
    cyan: 'FF06B6D4',
    green: 'FF16A34A',
    greenLight: 'FFDCFCE7',
    amber: 'FFD97706',
    amberLight: 'FFFEF3C7',
    red: 'FFDC2626',
    redLight: 'FFFEE2E2',
    purple: 'FF7C3AED',
    purpleLight: 'FFEDE9FE',
    slate50: 'FFF8FAFC',
    slate100: 'FFF1F5F9',
    slate200: 'FFE2E8F0',
    slate500: 'FF64748B',
    slate700: 'FF334155',
    white: 'FFFFFFFF'
  };

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

  function pct(
    value,
    digits = 1
  ) {
    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {
      return '';
    }

    const n =
      Number(value);

    if (
      Number.isNaN(n)
    ) {
      return '';
    }

    return `${(
      n * 100
    ).toFixed(digits)}%`;
  }

  function formatDate(
    timestamp
  ) {
    if (!timestamp) {
      return '';
    }

    return new Date(
      Number(timestamp) *
      1000
    ).toLocaleString(
      'zh-CN'
    );
  }

  function csvEscape(
    value
  ) {
    if (
      value === null ||
      value === undefined
    ) {
      return '';
    }

    const str =
      String(value);

    return /[",\n\r]/.test(str)
      ? `"${str.replace(
          /"/g,
          '""'
        )}"`
      : str;
  }

  function safeClone(
    value
  ) {
    try {
      return JSON.parse(
        JSON.stringify(
          value
        )
      );
    } catch {
      return null;
    }
  }

  function mean(
    values
  ) {
    const arr =
      values.filter(
        v =>
          typeof v ===
            'number' &&
          Number.isFinite(v)
      );

    if (!arr.length) {
      return '';
    }

    return (
      arr.reduce(
        (a, b) =>
          a + b,
        0
      ) /
      arr.length
    );
  }

  function median(
    values
  ) {
    const arr =
      values
        .filter(
          v =>
            typeof v ===
              'number' &&
            Number.isFinite(v)
        )
        .sort(
          (a, b) =>
            a - b
        );

    if (!arr.length) {
      return '';
    }

    const mid =
      Math.floor(
        arr.length / 2
      );

    return (
      arr.length % 2
        ? arr[mid]
        : (
            arr[mid - 1] +
            arr[mid]
          ) / 2
    );
  }

  function quantile(
    values,
    q
  ) {
    const arr =
      values
        .filter(
          v =>
            typeof v ===
              'number' &&
            Number.isFinite(v)
        )
        .sort(
          (a, b) =>
            a - b
        );

    if (!arr.length) {
      return '';
    }

    if (
      arr.length === 1
    ) {
      return arr[0];
    }

    const pos =
      (
        arr.length - 1
      ) * q;

    const base =
      Math.floor(pos);

    const rest =
      pos - base;

    return (
      arr[base + 1] !==
        undefined
        ? (
            arr[base] +
            rest *
            (
              arr[base + 1] -
              arr[base]
            )
          )
        : arr[base]
    );
  }

  function isScheduled(
    item
  ) {
    const now =
      Math.floor(
        Date.now() /
        1000
      );

    const postTime =
      Number(
        item?.post_time ||
        0
      );

    const scheduleTime =
      Number(
        item?.schedule_time ||
        0
      );

    return (
      postTime >
        now + 60 ||
      scheduleTime >
        now + 60
    );
  }

  function candidateItems() {
    return [
      ...items.values()
    ].filter(
      item =>
        !isScheduled(item)
    );
  }

  function scheduledItems() {
    return [
      ...items.values()
    ].filter(
      isScheduled
    );
  }

  function captureItemList(
    url,
    data
  ) {
    if (
      !url ||
      !String(url)
        .includes(
          '/tiktok/creator/manage/item_list/v1/'
        )
    ) {
      return;
    }

    if (
      !data ||
      !Array.isArray(
        data.item_list
      )
    ) {
      return;
    }

    for (
      const item
      of data.item_list
    ) {
      if (
        !item?.item_id
      ) {
        continue;
      }

      items.set(
        String(
          item.item_id
        ),
        safeClone(item)
      );
    }

    updateUI(
      data.has_more
        ? `已发现 ${items.size} 条，继续加载…`
        : `已发现 ${items.size} 条`
    );
  }

  page.fetch =
    async function (
      ...args
    ) {
      const response =
        await originalFetch(
          ...args
        );

      try {
        const input =
          args[0];

        const url =
          typeof input ===
          'string'
            ? input
            : input?.url;

        if (
          url &&
          String(url)
            .includes(
              '/tiktok/creator/manage/item_list/v1/'
            )
        ) {
          response
            .clone()
            .json()
            .then(
              data =>
                captureItemList(
                  url,
                  data
                )
            )
            .catch(
              () => {}
            );
        }
      } catch {}

      return response;
    };

  if (
    OriginalXHR
      ?.prototype
  ) {
    const originalOpen =
      OriginalXHR
        .prototype
        .open;

    const originalSend =
      OriginalXHR
        .prototype
        .send;

    OriginalXHR
      .prototype
      .open =
      function (
        method,
        url,
        ...rest
      ) {
        this
          .__qualitellTikTokURL =
          url;

        return originalOpen
          .call(
            this,
            method,
            url,
            ...rest
          );
      };

    OriginalXHR
      .prototype
      .send =
      function (
        ...args
      ) {
        this
          .addEventListener(
            'load',
            () => {
              try {
                const url =
                  this
                    .__qualitellTikTokURL;

                if (
                  url &&
                  String(url)
                    .includes(
                      '/tiktok/creator/manage/item_list/v1/'
                    )
                ) {
                  captureItemList(
                    url,
                    JSON.parse(
                      this
                        .responseText
                    )
                  );
                }
              } catch {}
            }
          );

        return originalSend
          .apply(
            this,
            args
          );
      };
  }

  function getCsrfToken() {
    const entry =
      document.cookie
        .split('; ')
        .find(
          cookie =>
            cookie
              .startsWith(
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

  async function fetchInsight(
    videoId
  ) {
    const typeRequests = [
      {
        insigh_type:
          'video_info',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_traffic_source_percent_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_retention_rate_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_view_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_total_duration_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_per_duration_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_finish_rate_realtime',
        aweme_id:
          videoId
      },
      {
        insigh_type:
          'video_new_follower_realtime',
        aweme_id:
          videoId
      }
    ];

    const lang =
      document
        .documentElement
        .lang ||
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
        locale:
          lang,

        aid:
          '1988',

        priority_region:
          'US',

        region:
          'US',

        tz_name:
          timezone,

        app_name:
          'tiktok_creator_center',

        app_language:
          lang,

        device_platform:
          'web_pc',

        channel:
          'tiktok_web',

        tz_offset:
          String(
            tzOffset
          ),

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
      headers[
        'tt-csrf-token'
      ] = csrf;
    }

    const response =
      await originalFetch(
        '/aweme/v2/data/insight/?' +
        params
          .toString(),
        {
          credentials:
            'include',

          headers
        }
      );

    return {
      httpStatus:
        response.status,

      data:
        await response
          .json()
    };
  }

  function retentionAt(
    list,
    milliseconds
  ) {
    const hit =
      (
        list || []
      ).find(
        item =>
          String(
            item.timestamp
          ) ===
          String(
            milliseconds
          )
      );

    return hit
      ? Number(
          hit.value
        )
      : '';
  }

  function trafficToObject(
    list
  ) {
    const result = {};

    for (
      const item
      of list || []
    ) {
      result[
        item.key
      ] =
        Number(
          item.value
        );
    }

    return result;
  }

  function extractUrl(
    value
  ) {
    if (!value) {
      return '';
    }

    if (
      typeof value ===
      'string'
    ) {
      return value;
    }

    if (
      Array.isArray(
        value
      )
    ) {
      return (
        value.find(
          v =>
            typeof v ===
            'string'
        ) ||
        ''
      );
    }

    if (
      Array.isArray(
        value.url_list
      )
    ) {
      return (
        value
          .url_list[0] ||
        ''
      );
    }

    if (
      typeof value.url ===
      'string'
    ) {
      return value.url;
    }

    return '';
  }

  function pickCoverUrl(
    item,
    videoInfo
  ) {
    const candidates = [
      item?.cover_url,
      item?.cover,
      item?.video?.cover,
      item?.video
        ?.origin_cover,
      item?.video
        ?.dynamic_cover,

      videoInfo?.cover,
      videoInfo?.video
        ?.cover,
      videoInfo?.video
        ?.origin_cover,
      videoInfo?.video
        ?.dynamic_cover
    ];

    for (
      const candidate
      of candidates
    ) {
      const url =
        extractUrl(
          candidate
        );

      if (url) {
        return url;
      }
    }

    return '';
  }

  function getVideoUrl(
    row
  ) {
    const account =
      String(
        row?.account ||
        ''
      ).replace(
        /^@/,
        ''
      );

    const videoId =
      String(
        row?.video_id ||
        ''
      );

    if (
      !account ||
      !videoId
    ) {
      return '';
    }

    return (
      `https://www.tiktok.com/@${account}/video/${videoId}`
    );
  }

  function normalizeRow(
    item,
    insight
  ) {
    const data =
      insight.data ||
      {};

    const videoInfo =
      data.video_info ||
      {};

    const statistics =
      videoInfo
        .statistics ||
      {};

    const retention =
      data
        .video_retention_rate_realtime
        ?.value
        ?.list ||
      [];

    const traffic =
      trafficToObject(
        data
          .video_traffic_source_percent_realtime
          ?.value
          ?.value ||
        []
      );

    const durationMs =
      Number(
        videoInfo.video
          ?.duration ||
        item.duration ||
        0
      );

    const durationSec =
      durationMs
        ? (
            durationMs /
            1000
          )
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
        statistics
          .play_count
      );

    const likes =
      num(
        statistics
          .digg_count ??
        item.like_count
      );

    const comments =
      num(
        statistics
          .comment_count ??
        item.comment_count
      );

    const shares =
      num(
        statistics
          .share_count ??
        item.share_count
      );

    const favorites =
      num(
        statistics
          .collect_count ??
        item.favorite_count
      );

    const publishTs =
      Number(
        item.post_time ||
        videoInfo
          .create_time ||
        0
      );

    const account =
      videoInfo
        .author
        ?.unique_id ||
      item.author
        ?.unique_id ||
      item
        .author_unique_id ||
      '';

    return {
      account,

      nickname:
        videoInfo
          .author
          ?.nickname ||
        item.author
          ?.nickname ||
        '',

      video_id:
        String(
          item.item_id ||
          videoInfo
            .aweme_id ||
          ''
        ),

      title:
        item.desc ||
        videoInfo.desc ||
        '',

      cover_url:
        pickCoverUrl(
          item,
          videoInfo
        ),

      publish_time:
        formatDate(
          publishTs
        ),

      publish_ts:
        publishTs,

      duration_sec:
        durationSec === ''
          ? ''
          : Number(
              durationSec
                .toFixed(3)
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
          averageWatch !==
            '' &&
          durationSec
        )
          ? (
              averageWatch /
              durationSec
            )
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
        traffic[
          'For You'
        ] ?? '',

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
          ? (
              likes /
              views
            )
          : '',

      engagement_rate:
        views
          ? (
              (
                likes +
                comments +
                shares +
                favorites
              ) /
              views
            )
          : ''
    };
  }

  function getRangeDays() {
    const value =
      $('qtk-range')
        ?.value ||
      '7';

    return (
      value === 'all'
        ? null
        : Number(
            value
          )
    );
  }

  function getAnalysisRows() {
    const rows = [
      ...detailRows
        .values()
    ];

    const days =
      getRangeDays();

    if (!days) {
      return rows
        .sort(
          (a, b) =>
            b.publish_ts -
            a.publish_ts
        );
    }

    const cutoff =
      Math.floor(
        Date.now() /
        1000
      ) -
      days *
      86400;

    return rows
      .filter(
        row =>
          row.publish_ts >=
          cutoff
      )
      .sort(
        (a, b) =>
          b.publish_ts -
          a.publish_ts
      );
  }

  function summaryFor(
    rows
  ) {
    return {
      count:
        rows.length,

      avgViews:
        mean(
          rows.map(
            r =>
              r.views
          )
        ),

      medianViews:
        median(
          rows.map(
            r =>
              r.views
          )
        ),

      avgFinish:
        mean(
          rows.map(
            r =>
              r.finish_rate
          )
        ),

      avgWatchRatio:
        mean(
          rows.map(
            r =>
              r.watch_ratio
          )
        ),

      avgR1:
        mean(
          rows.map(
            r =>
              r.retention_1s
          )
        ),

      avgR2:
        mean(
          rows.map(
            r =>
              r.retention_2s
          )
        ),

      avgR3:
        mean(
          rows.map(
            r =>
              r.retention_3s
          )
        ),

      avgR5:
        mean(
          rows.map(
            r =>
              r.retention_5s
          )
        )
    };
  }

  function getMetricValue(
    row,
    metric
  ) {
    return Number(
      row?.[metric]
    );
  }

  function topRows(
    rows,
    metric,
    limit = 5
  ) {
    return rows
      .filter(
        r =>
          Number.isFinite(
            getMetricValue(
              r,
              metric
            )
          )
      )
      .sort(
        (a, b) =>
          getMetricValue(
            b,
            metric
          ) -
          getMetricValue(
            a,
            metric
          )
      )
      .slice(
        0,
        limit
      );
  }

  function metricLabel(
    metric
  ) {
    const map = {
      views:
        '播放量',

      finish_rate:
        '完播率',

      watch_ratio:
        '观看倍率',

      retention_1s:
        '1秒留存',

      retention_3s:
        '3秒留存',

      new_followers:
        '新增粉丝',

      engagement_rate:
        '互动率'
    };

    return (
      map[metric] ||
      metric
    );
  }

  function metricDisplay(
    metric,
    value
  ) {
    if (
      !Number.isFinite(
        Number(value)
      )
    ) {
      return '-';
    }

    if (
      [
        'finish_rate',
        'watch_ratio',
        'retention_1s',
        'retention_3s',
        'engagement_rate'
      ].includes(
        metric
      )
    ) {
      return pct(
        Number(value)
      );
    }

    return String(
      value
    );
  }

  function buildSignals(
    rows
  ) {
    if (
      rows.length < 3
    ) {
      return {
        high: [],
        hidden: [],
        weak: []
      };
    }

    const views =
      rows
        .map(
          r =>
            r.views
        )
        .filter(
          Number.isFinite
        );

    const r3s =
      rows
        .map(
          r =>
            r.retention_3s
        )
        .filter(
          Number.isFinite
        );

    const q80Views =
      quantile(
        views,
        0.8
      );

    const medViews =
      median(
        views
      );

    const q75R3 =
      quantile(
        r3s,
        0.75
      );

    const q25R3 =
      quantile(
        r3s,
        0.25
      );

    return {
      high:
        rows
          .filter(
            r =>
              Number.isFinite(
                r.views
              ) &&
              r.views >=
                q80Views
          )
          .sort(
            (a, b) =>
              b.views -
              a.views
          )
          .slice(
            0,
            4
          ),

      hidden:
        rows
          .filter(
            r =>
              Number.isFinite(
                r.retention_3s
              ) &&
              Number.isFinite(
                r.views
              ) &&
              r.retention_3s >=
                q75R3 &&
              r.views <=
                medViews
          )
          .sort(
            (a, b) =>
              b.retention_3s -
              a.retention_3s
          )
          .slice(
            0,
            4
          ),

      weak:
        rows
          .filter(
            r =>
              Number.isFinite(
                r.retention_3s
              ) &&
              r.retention_3s <=
                q25R3
          )
          .sort(
            (a, b) =>
              a.retention_3s -
              b.retention_3s
          )
          .slice(
            0,
            4
          )
    };
  }

  function shortTitle(
    title,
    max = 38
  ) {
    const clean =
      String(
        title || ''
      )
        .replace(
          /\s+/g,
          ' '
        )
        .trim();

    return (
      clean.length >
        max
        ? (
            clean.slice(
              0,
              max
            ) +
            '…'
          )
        : clean
    );
  }

  function escapeHtml(
    str
  ) {
    return String(
      str ?? ''
    )
      .replace(
        /&/g,
        '&amp;'
      )
      .replace(
        /</g,
        '&lt;'
      )
      .replace(
        />/g,
        '&gt;'
      )
      .replace(
        /"/g,
        '&quot;'
      )
      .replace(
        /'/g,
        '&#039;'
      );
  }

  function renderTopList(
    rows,
    metric
  ) {
    if (
      !rows.length
    ) {
      return `
        <div class="qtk-empty">
          当前范围暂无已采集公开视频。
        </div>
      `;
    }

    const list =
      topRows(
        rows,
        metric,
        5
      );

    if (
      !list.length
    ) {
      return `
        <div class="qtk-empty">
          此指标暂无数据。
        </div>
      `;
    }

    return list
      .map(
        (r, i) => `
          <div class="qtk-top-item">

            <div class="qtk-rank">
              ${i + 1}
            </div>

            <div class="qtk-top-main">

              <div
                class="qtk-top-title"
                title="${escapeHtml(r.title)}"
              >
                ${escapeHtml(shortTitle(r.title))}
              </div>

              <div class="qtk-top-meta">
                播放 ${r.views || 0}
                · 完播 ${pct(r.finish_rate)}
                · 3秒 ${pct(r.retention_3s)}
              </div>

            </div>

            <div class="qtk-top-value">
              ${metricDisplay(metric, r[metric])}
            </div>

          </div>
        `
      )
      .join('');
  }

  function renderSignalBlock(
    title,
    icon,
    rows,
    mode
  ) {
    if (
      !rows.length
    ) {
      return '';
    }

    const content =
      rows
        .map(
          r => {
            let tag = '';

            if (
              mode ===
              'high'
            ) {
              tag =
                `播放 ${r.views}`;
            }

            if (
              mode ===
              'hidden'
            ) {
              tag =
                `3秒 ${pct(r.retention_3s)} · 播放 ${r.views}`;
            }

            if (
              mode ===
              'weak'
            ) {
              tag =
                `3秒 ${pct(r.retention_3s)}`;
            }

            return `
              <div class="qtk-signal-item">

                <span>
                  ${escapeHtml(shortTitle(r.title, 34))}
                </span>

                <b>
                  ${tag}
                </b>

              </div>
            `;
          }
        )
        .join('');

    return `
      <div class="qtk-signal-block">

        <div class="qtk-signal-head">
          ${icon} ${title}
        </div>

        ${content}

      </div>
    `;
  }

  function renderAnalysis() {
    if (
      !$(
        'qtk-analysis'
      )
    ) {
      return;
    }

    const rows =
      getAnalysisRows();

    const summary =
      summaryFor(
        rows
      );

    $('qtk-summary-count')
      .textContent =
      summary.count ||
      0;

    $('qtk-summary-avg')
      .textContent =
      summary.avgViews ===
        ''
        ? '-'
        : Math.round(
            summary.avgViews
          );

    $('qtk-summary-median')
      .textContent =
      summary.medianViews ===
        ''
        ? '-'
        : Math.round(
            summary.medianViews
          );

    $('qtk-summary-finish')
      .textContent =
      pct(
        summary.avgFinish
      );

    $('qtk-summary-watch')
      .textContent =
      pct(
        summary.avgWatchRatio
      );

    $('qtk-summary-r3')
      .textContent =
      pct(
        summary.avgR3
      );

    const metric =
      $('qtk-rank-metric')
        ?.value ||
      'views';

    $('qtk-top-list')
      .innerHTML =
      renderTopList(
        rows,
        metric
      );

    const signals =
      buildSignals(
        rows
      );

    $('qtk-signals')
      .innerHTML =
      (
        renderSignalBlock(
          '高播放样本（本期前20%）',
          '🔥',
          signals.high,
          'high'
        ) +
        renderSignalBlock(
          '高留存但播放一般',
          '👀',
          signals.hidden,
          'hidden'
        ) +
        renderSignalBlock(
          '3秒留存较弱（本期后25%）',
          '⚠️',
          signals.weak,
          'weak'
        )
      );
  }

  function injectStyle() {
    if (
      document
        .getElementById(
          'qualitell-tiktok-style'
        )
    ) {
      return;
    }

    const style =
      document
        .createElement(
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
        right:18px;
        bottom:18px;
        z-index:2147483647;
        width:390px;
        max-height:88vh;
        overflow:hidden;
        background:rgba(22,22,26,.985);
        color:#fff;
        border:1px solid rgba(255,255,255,.12);
        border-radius:14px;
        box-shadow:0 15px 45px rgba(0,0,0,.38);
        font-family:Arial,"Microsoft YaHei",sans-serif;
      }

      .qtk-header {
        display:flex;
        justify-content:space-between;
        align-items:center;
        padding:13px 15px;
        border-bottom:1px solid rgba(255,255,255,.08);
      }

      .qtk-title {
        font-size:15px;
        font-weight:700;
      }

      .qtk-version {
        margin-top:3px;
        font-size:11px;
        color:rgba(255,255,255,.52);
      }

      .qtk-minimize {
        border:0;
        background:transparent;
        color:#fff;
        font-size:20px;
        cursor:pointer;
      }

      #qtk-body {
        max-height:calc(88vh - 56px);
        overflow:auto;
        padding:13px 15px 16px;
      }

      .qtk-row {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:12px;
        padding:4px 0;
        font-size:12px;
      }

      .qtk-row span {
        color:rgba(255,255,255,.62);
      }

      .qtk-row b {
        text-align:right;
        font-weight:600;
      }

      .qtk-toolbar {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:7px;
        margin-top:10px;
      }

      .qtk-toolbar button,
      .qtk-select {
        width:100%;
        min-height:34px;
        border:0;
        border-radius:8px;
        padding:8px 9px;
        font-size:12px;
      }

      .qtk-toolbar button {
        cursor:pointer;
        font-weight:700;
      }

      #qtk-scan {
        background:#fe2c55;
        color:#fff;
      }

      #qtk-fetch {
        background:#25f4ee;
        color:#111;
      }

      #qtk-export-xlsx,
      #qtk-export-csv {
        background:#fff;
        color:#111;
      }

      #qtk-reset {
        background:#393940;
        color:#fff;
      }

      .qtk-toolbar button:disabled {
        opacity:.35;
        cursor:not-allowed;
      }

      .qtk-progress {
        height:7px;
        background:rgba(255,255,255,.1);
        border-radius:999px;
        margin-top:10px;
        overflow:hidden;
      }

      #qtk-progress-bar {
        height:100%;
        width:0%;
        background:#25f4ee;
        transition:width .2s ease;
      }

      .qtk-divider {
        height:1px;
        background:rgba(255,255,255,.08);
        margin:13px 0;
      }

      .qtk-section-title {
        display:flex;
        justify-content:space-between;
        align-items:center;
        font-size:13px;
        font-weight:700;
        margin-bottom:8px;
      }

      .qtk-select {
        background:#303036;
        color:#fff;
        outline:none;
      }

      .qtk-summary-grid {
        display:grid;
        grid-template-columns:repeat(3,1fr);
        gap:7px;
      }

      .qtk-card {
        background:#2a2a30;
        border:1px solid rgba(255,255,255,.06);
        border-radius:9px;
        padding:9px;
      }

      .qtk-card span {
        display:block;
        font-size:10px;
        color:rgba(255,255,255,.5);
        margin-bottom:5px;
      }

      .qtk-card b {
        font-size:15px;
        font-weight:700;
      }

      .qtk-rank-controls {
        display:grid;
        grid-template-columns:1fr;
        gap:7px;
        margin:10px 0 7px;
      }

      .qtk-top-item {
        display:flex;
        align-items:center;
        gap:8px;
        padding:7px 0;
        border-bottom:1px solid rgba(255,255,255,.06);
      }

      .qtk-rank {
        width:20px;
        height:20px;
        border-radius:6px;
        background:#34343b;
        display:flex;
        align-items:center;
        justify-content:center;
        font-size:11px;
        font-weight:700;
        flex:none;
      }

      .qtk-top-main {
        min-width:0;
        flex:1;
      }

      .qtk-top-title {
        font-size:11px;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
      }

      .qtk-top-meta {
        margin-top:3px;
        font-size:9px;
        color:rgba(255,255,255,.45);
      }

      .qtk-top-value {
        font-size:11px;
        font-weight:700;
        color:#25f4ee;
        flex:none;
      }

      .qtk-signal-block {
        margin-top:9px;
        padding:9px;
        background:#26262c;
        border-radius:9px;
      }

      .qtk-signal-head {
        font-size:11px;
        font-weight:700;
        margin-bottom:6px;
      }

      .qtk-signal-item {
        display:flex;
        justify-content:space-between;
        gap:10px;
        padding:4px 0;
        font-size:10px;
      }

      .qtk-signal-item span {
        overflow:hidden;
        text-overflow:ellipsis;
        white-space:nowrap;
      }

      .qtk-signal-item b {
        flex:none;
        font-weight:600;
        color:rgba(255,255,255,.65);
      }

      .qtk-empty {
        font-size:11px;
        color:rgba(255,255,255,.45);
        padding:8px 0;
      }

      .qtk-tip {
        margin-top:10px;
        font-size:10px;
        line-height:1.5;
        color:rgba(255,255,255,.42);
      }
    `;

    (
      document.head ||
      document
        .documentElement
    )
      .appendChild(
        style
      );
  }

  function buildPanel() {
    if (
      !document
        .documentElement ||
      document
        .getElementById(
          'qualitell-tiktok-panel'
        )
    ) {
      return;
    }

    injectStyle();

    const panel =
      document
        .createElement(
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
            V${VERSION} · 美化 Excel + 视频封面
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

        <div class="qtk-toolbar">

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
            id="qtk-export-xlsx"
            disabled
          >
            导出美化 XLSX
          </button>

          <button
            id="qtk-export-csv"
            disabled
          >
            导出 CSV
          </button>

          <button
            id="qtk-reset"
            style="grid-column:1/3"
          >
            清空重扫
          </button>

        </div>

        <div class="qtk-progress">
          <div id="qtk-progress-bar"></div>
        </div>

        <div class="qtk-divider"></div>

        <div id="qtk-analysis">

          <div class="qtk-section-title">

            <span>
              本期分析
            </span>

            <select
              id="qtk-range"
              class="qtk-select"
              style="width:120px"
            >
              <option value="7">
                最近7天
              </option>

              <option value="14">
                最近14天
              </option>

              <option value="30">
                最近30天
              </option>

              <option value="all">
                全部
              </option>
            </select>

          </div>

          <div class="qtk-summary-grid">

            <div class="qtk-card">
              <span>公开视频</span>
              <b id="qtk-summary-count">0</b>
            </div>

            <div class="qtk-card">
              <span>平均播放</span>
              <b id="qtk-summary-avg">-</b>
            </div>

            <div class="qtk-card">
              <span>中位播放</span>
              <b id="qtk-summary-median">-</b>
            </div>

            <div class="qtk-card">
              <span>平均完播</span>
              <b id="qtk-summary-finish">-</b>
            </div>

            <div class="qtk-card">
              <span>观看倍率</span>
              <b id="qtk-summary-watch">-</b>
            </div>

            <div class="qtk-card">
              <span>3秒留存</span>
              <b id="qtk-summary-r3">-</b>
            </div>

          </div>

          <div class="qtk-rank-controls">

            <select
              id="qtk-rank-metric"
              class="qtk-select"
            >
              <option value="views">
                播放 TOP
              </option>

              <option value="finish_rate">
                完播率 TOP
              </option>

              <option value="watch_ratio">
                观看倍率 TOP
              </option>

              <option value="retention_1s">
                1秒留存 TOP
              </option>

              <option value="retention_3s">
                3秒留存 TOP
              </option>

              <option value="new_followers">
                新增粉丝 TOP
              </option>

              <option value="engagement_rate">
                互动率 TOP
              </option>
            </select>

          </div>

          <div id="qtk-top-list">
            <div class="qtk-empty">
              采集数据后显示排名。
            </div>
          </div>

          <div id="qtk-signals"></div>

        </div>

        <div class="qtk-tip">
          XLSX 将自动加入小尺寸视频封面；标题和“打开视频”均可直接点击进入 TikTok 视频。
        </div>

      </div>
    `;

    document
      .documentElement
      .appendChild(
        panel
      );

    $('qtk-minimize')
      .addEventListener(
        'click',
        () => {
          minimized =
            !minimized;

          $('qtk-body')
            .style
            .display =
            minimized
              ? 'none'
              : '';

          $('qtk-minimize')
            .textContent =
            minimized
              ? '+'
              : '−';
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

    $('qtk-export-xlsx')
      .addEventListener(
        'click',
        exportXlsx
      );

    $('qtk-export-csv')
      .addEventListener(
        'click',
        exportCsv
      );

    $('qtk-reset')
      .addEventListener(
        'click',
        resetAll
      );

    $('qtk-range')
      .addEventListener(
        'change',
        renderAnalysis
      );

    $('qtk-rank-metric')
      .addEventListener(
        'change',
        renderAnalysis
      );

    updateUI(
      '已就绪'
    );

    renderAnalysis();
  }

  function updateUI(
    status
  ) {
    if (
      !$(
        'qualitell-tiktok-panel'
      )
    ) {
      return;
    }

    $('qtk-all')
      .textContent =
      items.size;

    $('qtk-scheduled')
      .textContent =
      scheduledItems()
        .length;

    $('qtk-candidates')
      .textContent =
      candidateItems()
        .length;

    $('qtk-public')
      .textContent =
      detailRows.size;

    $('qtk-private')
      .textContent =
      privateIds.size;

    $('qtk-failed')
      .textContent =
      failedCount +
      unknownIds.size;

    if (status) {
      $('qtk-status')
        .textContent =
        status;
    }

    $('qtk-scan')
      .disabled =
      busy;

    $('qtk-fetch')
      .disabled =
      busy ||
      candidateItems()
        .length ===
        0;

    $('qtk-export-xlsx')
      .disabled =
      busy ||
      detailRows.size ===
        0;

    $('qtk-export-csv')
      .disabled =
      busy ||
      detailRows.size ===
        0;

    renderAnalysis();
  }

  async function scanAll() {
    if (busy) {
      return;
    }

    busy = true;

    updateUI(
      '正在扫描作品…'
    );

    let stableRounds =
      0;

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

      await sleep(
        850
      );

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
        stableRounds >=
        6
      ) {
        break;
      }
    }

    busy = false;

    updateUI(
      items.size
        ? `扫描完成：${items.size} 条`
        : '未发现作品，请刷新作品页'
    );
  }

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
      .style
      .width =
      '0%';

    updateUI(
      '开始获取详细数据…'
    );

    for (
      const item
      of list
    ) {
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
            ?.status_code !==
            0
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

        if (
          privateStatus ===
          0
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

      } catch (
        error
      ) {
        console.warn(
          '[TikTok Collector]',
          videoId,
          error
        );

        failedCount++;
      }

      completed++;

      $('qtk-progress-bar')
        .style
        .width =
        `${
          list.length
            ? Math.round(
                completed /
                list.length *
                100
              )
            : 0
        }%`;

      updateUI(
        `详情 ${completed}/${list.length}`
      );

      await sleep(
        800
      );
    }

    busy = false;

    updateUI(
      `完成：公开 ${detailRows.size}，非公开 ${privateIds.size}，失败/未知 ${failedCount + unknownIds.size}`
    );
  }

  function resetAll() {
    if (busy) {
      return;
    }

    items.clear();
    detailRows.clear();
    privateIds.clear();
    unknownIds.clear();
    coverCache.clear();

    failedCount = 0;

    if (
      $('qtk-progress-bar')
    ) {
      $('qtk-progress-bar')
        .style
        .width =
        '0%';
    }

    updateUI(
      '已清空，请刷新页面重新扫描'
    );
  }

  const exportColumns = [
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

  function exportCsv() {
    const rows = [
      ...detailRows
        .values()
    ];

    if (
      !rows.length
    ) {
      return;
    }

    const columns = [
      ...exportColumns,

      [
        '封面链接',
        '__cover_url'
      ],

      [
        '视频链接',
        '__video_url'
      ]
    ];

    const lines = [
      columns
        .map(
          ([label]) =>
            csvEscape(
              label
            )
        )
        .join(',')
    ];

    for (
      const row
      of rows
    ) {
      lines.push(
        columns
          .map(
            ([, key]) => {
              if (
                key ===
                '__cover_url'
              ) {
                return csvEscape(
                  row.cover_url ||
                  ''
                );
              }

              if (
                key ===
                '__video_url'
              ) {
                return csvEscape(
                  getVideoUrl(
                    row
                  )
                );
              }

              return csvEscape(
                percentKeys.has(
                  key
                )
                  ? pct(
                      row[key]
                    )
                  : row[key]
              );
            }
          )
          .join(',')
      );
    }

    const blob =
      new Blob(
        [
          '\uFEFF' +
          lines.join(
            '\r\n'
          )
        ],
        {
          type:
            'text/csv;charset=utf-8'
        }
      );

    const account =
      rows[0]
        ?.account ||
      'tiktok';

    const date =
      new Date()
        .toISOString()
        .slice(
          0,
          10
        );

    downloadBlob(
      blob,
      `${account}_TikTok公开视频数据_${date}.csv`
    );
  }

  function downloadBlob(
    blob,
    filename
  ) {
    const url =
      URL.createObjectURL(
        blob
      );

    const link =
      document
        .createElement(
          'a'
        );

    link.href =
      url;

    link.download =
      filename;

    document.body
      .appendChild(
        link
      );

    link.click();

    link.remove();

    setTimeout(
      () =>
        URL.revokeObjectURL(
          url
        ),
      3000
    );
  }

  function gmFetchBlob(
    url
  ) {
    return new Promise(
      (
        resolve,
        reject
      ) => {
        if (
          typeof GM_xmlhttpRequest !==
          'function'
        ) {
          reject(
            new Error(
              'GM_xmlhttpRequest unavailable'
            )
          );

          return;
        }

        GM_xmlhttpRequest({
          method:
            'GET',

          url,

          responseType:
            'arraybuffer',

          timeout:
            15000,

          onload:
            response => {
              if (
                response.status <
                  200 ||
                response.status >=
                  300 ||
                !response.response
              ) {
                reject(
                  new Error(
                    `cover http ${response.status}`
                  )
                );

                return;
              }

              const match =
                String(
                  response
                    .responseHeaders ||
                  ''
                ).match(
                  /content-type:\s*([^;\r\n]+)/i
                );

              const type =
                match
                  ?.[1]
                  ?.trim() ||
                'image/jpeg';

              resolve(
                new Blob(
                  [
                    response
                      .response
                  ],
                  {
                    type
                  }
                )
              );
            },

          ontimeout:
            () =>
              reject(
                new Error(
                  'cover timeout'
                )
              ),

          onerror:
            () =>
              reject(
                new Error(
                  'cover network error'
                )
              )
        });
      }
    );
  }

  async function fetchCoverBlob(
    url
  ) {
    try {
      return await gmFetchBlob(
        url
      );
    } catch {
      const response =
        await originalFetch(
          url,
          {
            credentials:
              'omit',

            cache:
              'force-cache'
          }
        );

      if (
        !response.ok
      ) {
        throw new Error(
          `cover http ${response.status}`
        );
      }

      return await response
        .blob();
    }
  }

  async function blobToThumbnailDataUrl(
    blob,
    maxWidth = 54,
    maxHeight = 96
  ) {
    let source;
    let width;
    let height;

    let cleanup =
      () => {};

    if (
      typeof createImageBitmap ===
      'function'
    ) {
      source =
        await createImageBitmap(
          blob
        );

      width =
        source.width;

      height =
        source.height;

      cleanup =
        () =>
          source.close
            ?.();

    } else {
      const objectUrl =
        URL.createObjectURL(
          blob
        );

      const img =
        new Image();

      await new Promise(
        (
          resolve,
          reject
        ) => {
          img.onload =
            resolve;

          img.onerror =
            reject;

          img.src =
            objectUrl;
        }
      );

      source =
        img;

      width =
        img.naturalWidth ||
        img.width;

      height =
        img.naturalHeight ||
        img.height;

      cleanup =
        () =>
          URL.revokeObjectURL(
            objectUrl
          );
    }

    try {
      const scale =
        Math.min(
          maxWidth /
            width,

          maxHeight /
            height,

          1
        );

      const canvas =
        document
          .createElement(
            'canvas'
          );

      canvas.width =
        Math.max(
          1,
          Math.round(
            width *
            scale
          )
        );

      canvas.height =
        Math.max(
          1,
          Math.round(
            height *
            scale
          )
        );

      const ctx =
        canvas.getContext(
          '2d'
        );

      ctx.fillStyle =
        '#ffffff';

      ctx.fillRect(
        0,
        0,
        canvas.width,
        canvas.height
      );

      ctx.drawImage(
        source,
        0,
        0,
        canvas.width,
        canvas.height
      );

      return canvas
        .toDataURL(
          'image/jpeg',
          0.72
        );

    } finally {
      cleanup();
    }
  }

  async function getCoverData(
    row
  ) {
    const url =
      row?.cover_url ||
      '';

    if (!url) {
      return '';
    }

    if (
      coverCache.has(
        url
      )
    ) {
      return coverCache.get(
        url
      );
    }

    const promise =
      (
        async () => {
          try {
            const blob =
              await fetchCoverBlob(
                url
              );

            return await blobToThumbnailDataUrl(
              blob
            );

          } catch (
            error
          ) {
            console.warn(
              '[TikTok Collector] cover failed',
              row.video_id,
              error
            );

            return '';
          }
        }
      )();

    coverCache.set(
      url,
      promise
    );

    return promise;
  }

  async function prepareCoverImages(
    rows,
    workbook
  ) {
    const uniqueRows =
      [];

    const seen =
      new Set();

    for (
      const row
      of rows
    ) {
      if (
        !row?.video_id ||
        seen.has(
          row.video_id
        )
      ) {
        continue;
      }

      seen.add(
        row.video_id
      );

      uniqueRows.push(
        row
      );
    }

    const imageIds =
      new Map();

    let index = 0;
    let done = 0;

    const workerCount =
      Math.min(
        4,
        uniqueRows.length ||
        1
      );

    async function worker() {
      while (true) {
        const current =
          index++;

        if (
          current >=
          uniqueRows.length
        ) {
          break;
        }

        const row =
          uniqueRows[
            current
          ];

        const dataUrl =
          await getCoverData(
            row
          );

        if (dataUrl) {
          try {
            const imageId =
              workbook.addImage({
                base64:
                  dataUrl,

                extension:
                  'jpeg'
              });

            imageIds.set(
              row.video_id,
              imageId
            );

          } catch (
            error
          ) {
            console.warn(
              '[TikTok Collector] add image failed',
              row.video_id,
              error
            );
          }
        }

        done++;

        const progress =
          uniqueRows.length
            ? Math.round(
                done /
                uniqueRows.length *
                100
              )
            : 100;

        $('qtk-progress-bar')
          .style
          .width =
          `${progress}%`;

        updateUI(
          `准备封面 ${done}/${uniqueRows.length}`
        );
      }
    }

    await Promise.all(
      Array.from(
        {
          length:
            workerCount
        },
        worker
      )
    );

    return imageIds;
  }

  function thinBorder(
    color =
      COLORS.slate200
  ) {
    const side = {
      style:
        'thin',

      color: {
        argb:
          color
      }
    };

    return {
      top:
        side,

      left:
        side,

      bottom:
        side,

      right:
        side
    };
  }

  function styleHeaderRow(
    row,
    fill =
      COLORS.navy
  ) {
    row.height = 26;

    row.eachCell(
      {
        includeEmpty:
          true
      },
      cell => {
        cell.font = {
          bold:
            true,

          color: {
            argb:
              COLORS.white
          },

          size:
            10.5
        };

        cell.fill = {
          type:
            'pattern',

          pattern:
            'solid',

          fgColor: {
            argb:
              fill
          }
        };

        cell.alignment = {
          vertical:
            'middle',

          horizontal:
            'center',

          wrapText:
            true
        };

        cell.border =
          thinBorder(
            fill
          );
      }
    );
  }

  function styleDataRow(
    row,
    zebra =
      false
  ) {
    row.eachCell(
      {
        includeEmpty:
          true
      },
      cell => {
        cell.font = {
          color: {
            argb:
              COLORS.slate700
          },

          size:
            10
        };

        cell.alignment = {
          vertical:
            'middle',

          wrapText:
            true
        };

        cell.border =
          thinBorder();

        if (zebra) {
          cell.fill = {
            type:
              'pattern',

            pattern:
              'solid',

            fgColor: {
              argb:
                COLORS.slate50
            }
          };
        }
      }
    );
  }

  function setLinkCell(
    cell,
    text,
    url
  ) {
    if (!url) {
      cell.value =
        text || '';

      return;
    }

    cell.value = {
      text:
        text ||
        '打开视频',

      hyperlink:
        url,

      tooltip:
        '打开 TikTok 视频'
    };

    cell.font = {
      color: {
        argb:
          COLORS.blue
      },

      underline:
        true,

      size:
        10
    };
  }

  function addCoverToRow(
    sheet,
    workbookImageIds,
    rowData,
    excelRowNumber,
    col = 1
  ) {
    const imageId =
      workbookImageIds
        .get(
          rowData.video_id
        );

    if (
      imageId ===
      undefined
    ) {
      return;
    }

    sheet.addImage(
      imageId,
      {
        tl: {
          col:
            col -
            1 +
            0.18,

          row:
            excelRowNumber -
            1 +
            0.08
        },

        ext: {
          width:
            42,

          height:
            74
        },

        editAs:
          'oneCell'
      }
    );
  }

  function buildRawSheet(
    workbook,
    rows,
    imageIds
  ) {
    const sheet =
      workbook
        .addWorksheet(
          '原始数据',
          {
            views: [
              {
                state:
                  'frozen',

                xSplit:
                  2,

                ySplit:
                  1
              }
            ],

            properties: {
              defaultRowHeight:
                20
            }
          }
        );

    sheet.columns = [
      {
        header:
          '封面',
        key:
          'cover',
        width:
          8
      },
      {
        header:
          '标题',
        key:
          'title',
        width:
          46
      },
      {
        header:
          '账号',
        key:
          'account',
        width:
          18
      },
      {
        header:
          '昵称',
        key:
          'nickname',
        width:
          18
      },
      {
        header:
          '发布时间',
        key:
          'publish_time',
        width:
          20
      },
      {
        header:
          '时长(s)',
        key:
          'duration_sec',
        width:
          11
      },
      {
        header:
          '播放量',
        key:
          'views',
        width:
          11
      },
      {
        header:
          '点赞',
        key:
          'likes',
        width:
          9
      },
      {
        header:
          '评论',
        key:
          'comments',
        width:
          9
      },
      {
        header:
          '分享',
        key:
          'shares',
        width:
          9
      },
      {
        header:
          '收藏',
        key:
          'favorites',
        width:
          9
      },
      {
        header:
          '新增粉丝',
        key:
          'new_followers',
        width:
          11
      },
      {
        header:
          '平均观看(s)',
        key:
          'avg_watch_sec',
        width:
          13
      },
      {
        header:
          '观看倍率',
        key:
          'watch_ratio',
        width:
          11
      },
      {
        header:
          '完播率',
        key:
          'finish_rate',
        width:
          11
      },
      {
        header:
          '总观看(s)',
        key:
          'total_watch_sec',
        width:
          13
      },
      {
        header:
          '1秒留存',
        key:
          'retention_1s',
        width:
          11
      },
      {
        header:
          '2秒留存',
        key:
          'retention_2s',
        width:
          11
      },
      {
        header:
          '3秒留存',
        key:
          'retention_3s',
        width:
          11
      },
      {
        header:
          '5秒留存',
        key:
          'retention_5s',
        width:
          11
      },
      {
        header:
          '10秒留存',
        key:
          'retention_10s',
        width:
          11
      },
      {
        header:
          'For You',
        key:
          'for_you',
        width:
          11
      },
      {
        header:
          '个人主页',
        key:
          'personal_profile',
        width:
          11
      },
      {
        header:
          '搜索',
        key:
          'search',
        width:
          10
      },
      {
        header:
          '关注',
        key:
          'follow',
        width:
          10
      },
      {
        header:
          '私信',
        key:
          'direct_message',
        width:
          10
      },
      {
        header:
          '音乐',
        key:
          'sound',
        width:
          10
      },
      {
        header:
          '其它',
        key:
          'others',
        width:
          10
      },
      {
        header:
          '点赞率',
        key:
          'like_rate',
        width:
          10
      },
      {
        header:
          '互动率',
        key:
          'engagement_rate',
        width:
          10
      },
      {
        header:
          'Video ID',
        key:
          'video_id',
        width:
          22
      },
      {
        header:
          '视频',
        key:
          'video_link',
        width:
          13
      }
    ];

    styleHeaderRow(
      sheet.getRow(1),
      COLORS.navy
    );

    sheet.autoFilter = {
      from: {
        row:
          1,

        column:
          1
      },

      to: {
        row:
          1,

        column:
          sheet.columnCount
      }
    };

    rows.forEach(
      (
        r,
        i
      ) => {
        const videoUrl =
          getVideoUrl(
            r
          );

        const row =
          sheet.addRow({
            cover:
              '',

            title:
              r.title ||
              '(无标题)',

            account:
              r.account,

            nickname:
              r.nickname,

            publish_time:
              r.publish_time,

            duration_sec:
              r.duration_sec,

            views:
              r.views,

            likes:
              r.likes,

            comments:
              r.comments,

            shares:
              r.shares,

            favorites:
              r.favorites,

            new_followers:
              r.new_followers,

            avg_watch_sec:
              r.avg_watch_sec,

            watch_ratio:
              r.watch_ratio,

            finish_rate:
              r.finish_rate,

            total_watch_sec:
              r.total_watch_sec,

            retention_1s:
              r.retention_1s,

            retention_2s:
              r.retention_2s,

            retention_3s:
              r.retention_3s,

            retention_5s:
              r.retention_5s,

            retention_10s:
              r.retention_10s,

            for_you:
              r.for_you,

            personal_profile:
              r.personal_profile,

            search:
              r.search,

            follow:
              r.follow,

            direct_message:
              r.direct_message,

            sound:
              r.sound,

            others:
              r.others,

            like_rate:
              r.like_rate,

            engagement_rate:
              r.engagement_rate,

            video_id:
              r.video_id,

            video_link:
              '打开视频'
          });

        row.height =
          58;

        styleDataRow(
          row,
          i % 2 === 1
        );

        setLinkCell(
          row.getCell(
            'title'
          ),
          r.title ||
            '(无标题)',
          videoUrl
        );

        setLinkCell(
          row.getCell(
            'video_link'
          ),
          '打开视频',
          videoUrl
        );

        addCoverToRow(
          sheet,
          imageIds,
          r,
          row.number,
          1
        );

        row
          .getCell(
            'cover'
          )
          .alignment = {
            vertical:
              'middle',

            horizontal:
              'center'
          };

        row
          .getCell(
            'title'
          )
          .alignment = {
            vertical:
              'middle',

            wrapText:
              true
          };
      }
    );

    [
      'views',
      'likes',
      'comments',
      'shares',
      'favorites',
      'new_followers',
      'total_watch_sec'
    ].forEach(
      key =>
        sheet
          .getColumn(
            key
          )
          .numFmt =
          '#,##0'
    );

    [
      'duration_sec',
      'avg_watch_sec'
    ].forEach(
      key =>
        sheet
          .getColumn(
            key
          )
          .numFmt =
          '0.00'
    );

    [
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
    ].forEach(
      key =>
        sheet
          .getColumn(
            key
          )
          .numFmt =
          '0.0%'
    );

    return sheet;
  }

  function styleSectionTitle(
    sheet,
    rowNumber,
    title,
    fillColor
  ) {
    sheet.mergeCells(
      rowNumber,
      1,
      rowNumber,
      8
    );

    const cell =
      sheet.getCell(
        rowNumber,
        1
      );

    cell.value =
      title;

    cell.font = {
      bold:
        true,

      color: {
        argb:
          COLORS.white
      },

      size:
        11
    };

    cell.fill = {
      type:
        'pattern',

      pattern:
        'solid',

      fgColor: {
        argb:
          fillColor
      }
    };

    cell.alignment = {
      vertical:
        'middle',

      horizontal:
        'left'
    };

    cell.border =
      thinBorder(
        fillColor
      );

    sheet
      .getRow(
        rowNumber
      )
      .height =
      24;
  }

  function styleAnalysisTableHeader(
    sheet,
    rowNumber
  ) {
    const values = [
      '封面',
      '视频标题',
      '当前指标',
      '播放量',
      '完播率',
      '3秒留存',
      '发布时间',
      '视频'
    ];

    const row =
      sheet.getRow(
        rowNumber
      );

    values.forEach(
      (
        value,
        i
      ) => {
        row
          .getCell(
            i + 1
          )
          .value =
          value;
      }
    );

    styleHeaderRow(
      row,
      COLORS.navy2
    );
  }

  function addAnalysisVideoRow(
    sheet,
    rowData,
    excelRowNumber,
    imageIds,
    metricKey,
    metricText
  ) {
    const row =
      sheet.getRow(
        excelRowNumber
      );

    const videoUrl =
      getVideoUrl(
        rowData
      );

    row.height =
      58;

    row
      .getCell(1)
      .value =
      '';

    setLinkCell(
      row.getCell(2),
      rowData.title ||
        '(无标题)',
      videoUrl
    );

    row
      .getCell(3)
      .value =
      rowData[
        metricKey
      ] ?? '';

    row
      .getCell(4)
      .value =
      rowData.views;

    row
      .getCell(5)
      .value =
      rowData
        .finish_rate;

    row
      .getCell(6)
      .value =
      rowData
        .retention_3s;

    row
      .getCell(7)
      .value =
      rowData
        .publish_time;

    setLinkCell(
      row.getCell(8),
      '打开视频',
      videoUrl
    );

    styleDataRow(
      row,
      excelRowNumber %
        2 ===
        0
    );

    row
      .getCell(2)
      .alignment = {
        vertical:
          'middle',

        wrapText:
          true
      };

    row
      .getCell(3)
      .numFmt =
      [
        'finish_rate',
        'watch_ratio',
        'retention_1s',
        'retention_3s',
        'engagement_rate'
      ].includes(
        metricKey
      )
        ? '0.0%'
        : '#,##0';

    row
      .getCell(4)
      .numFmt =
      '#,##0';

    row
      .getCell(5)
      .numFmt =
      '0.0%';

    row
      .getCell(6)
      .numFmt =
      '0.0%';

    if (metricText) {
      row
        .getCell(3)
        .note =
        metricText;
    }

    addCoverToRow(
      sheet,
      imageIds,
      rowData,
      excelRowNumber,
      1
    );
  }

  function buildAnalysisSheet(
    workbook,
    rows,
    imageIds
  ) {
    const sheet =
      workbook
        .addWorksheet(
          '本期分析',
          {
            views: [
              {
                state:
                  'frozen',

                ySplit:
                  7
              }
            ],

            properties: {
              defaultRowHeight:
                20
            }
          }
        );

    sheet.columns = [
      {
        width:
          8
      },
      {
        width:
          48
      },
      {
        width:
          15
      },
      {
        width:
          12
      },
      {
        width:
          12
      },
      {
        width:
          12
      },
      {
        width:
          20
      },
      {
        width:
          13
      }
    ];

    const summary =
      summaryFor(
        rows
      );

    const metric =
      $('qtk-rank-metric')
        ?.value ||
      'views';

    const rangeLabel =
      $('qtk-range')
        ?.selectedOptions
        ?.[0]
        ?.textContent
        ?.trim() ||
      '最近7天';

    const top =
      topRows(
        rows,
        metric,
        5
      );

    const signals =
      buildSignals(
        rows
      );

    sheet.mergeCells(
      'A1:H1'
    );

    const titleCell =
      sheet.getCell(
        'A1'
      );

    titleCell.value =
      'TikTok 内容表现分析';

    titleCell.font = {
      bold:
        true,

      size:
        18,

      color: {
        argb:
          COLORS.white
      }
    };

    titleCell.fill = {
      type:
        'pattern',

      pattern:
        'solid',

      fgColor: {
        argb:
          COLORS.navy
      }
    };

    titleCell.alignment = {
      vertical:
        'middle',

      horizontal:
        'left'
    };

    sheet
      .getRow(1)
      .height =
      32;

    sheet.mergeCells(
      'A2:H2'
    );

    const subCell =
      sheet.getCell(
        'A2'
      );

    subCell.value =
      `${rangeLabel} · 生成时间 ${new Date().toLocaleString('zh-CN')}`;

    subCell.font = {
      size:
        10,

      color: {
        argb:
          COLORS.slate500
      }
    };

    subCell.fill = {
      type:
        'pattern',

      pattern:
        'solid',

      fgColor: {
        argb:
          COLORS.slate100
      }
    };

    subCell.alignment = {
      vertical:
        'middle',

      horizontal:
        'left'
    };

    sheet
      .getRow(2)
      .height =
      22;

    const kpis = [
      [
        '公开视频',
        summary.count,
        '0'
      ],

      [
        '平均播放',
        summary.avgViews ===
          ''
          ? ''
          : Math.round(
              summary.avgViews
            ),
        '#,##0'
      ],

      [
        '中位播放',
        summary.medianViews ===
          ''
          ? ''
          : Math.round(
              summary.medianViews
            ),
        '#,##0'
      ],

      [
        '平均完播',
        summary.avgFinish,
        '0.0%'
      ],

      [
        '观看倍率',
        summary.avgWatchRatio,
        '0.0%'
      ],

      [
        '3秒留存',
        summary.avgR3,
        '0.0%'
      ]
    ];

    kpis.forEach(
      (
        [
          label,
          value,
          numFmt
        ],
        i
      ) => {
        const col =
          i + 1;

        const labelCell =
          sheet.getCell(
            4,
            col
          );

        const valueCell =
          sheet.getCell(
            5,
            col
          );

        labelCell.value =
          label;

        labelCell.font = {
          bold:
            true,

          color: {
            argb:
              COLORS.white
          },

          size:
            10
        };

        labelCell.fill = {
          type:
            'pattern',

          pattern:
            'solid',

          fgColor: {
            argb:
              COLORS.blue
          }
        };

        labelCell.alignment = {
          vertical:
            'middle',

          horizontal:
            'center'
        };

        labelCell.border =
          thinBorder(
            COLORS.blue
          );

        valueCell.value =
          value;

        valueCell.font = {
          bold:
            true,

          color: {
            argb:
              COLORS.navy
          },

          size:
            14
        };

        valueCell.fill = {
          type:
            'pattern',

          pattern:
            'solid',

          fgColor: {
            argb:
              COLORS.blueLight
          }
        };

        valueCell.alignment = {
          vertical:
            'middle',

          horizontal:
            'center'
        };

        valueCell.border =
          thinBorder(
            COLORS.slate200
          );

        valueCell.numFmt =
          numFmt;
      }
    );

    sheet
      .getRow(4)
      .height =
      22;

    sheet
      .getRow(5)
      .height =
      30;

    let rowNum =
      7;

    const addSection =
      (
        title,
        fillColor,
        sectionRows,
        metricKey,
        metricText
      ) => {
        styleSectionTitle(
          sheet,
          rowNum,
          title,
          fillColor
        );

        rowNum++;

        styleAnalysisTableHeader(
          sheet,
          rowNum
        );

        rowNum++;

        if (
          !sectionRows.length
        ) {
          sheet.mergeCells(
            rowNum,
            1,
            rowNum,
            8
          );

          const cell =
            sheet.getCell(
              rowNum,
              1
            );

          cell.value =
            '暂无符合条件的视频';

          cell.font = {
            italic:
              true,

            color: {
              argb:
                COLORS.slate500
            }
          };

          cell.fill = {
            type:
              'pattern',

            pattern:
              'solid',

            fgColor: {
              argb:
                COLORS.slate50
            }
          };

          cell.alignment = {
            vertical:
              'middle',

            horizontal:
              'center'
          };

          cell.border =
            thinBorder();

          rowNum++;

        } else {
          for (
            const item
            of sectionRows
          ) {
            addAnalysisVideoRow(
              sheet,
              item,
              rowNum,
              imageIds,
              metricKey,
              metricText
            );

            rowNum++;
          }
        }

        rowNum++;
      };

    addSection(
      `${metricLabel(metric)} TOP 5`,
      COLORS.blue,
      top,
      metric,
      metricLabel(metric)
    );

    addSection(
      '🔥 高播放样本（本期前20%）',
      COLORS.green,
      signals.high,
      'views',
      '播放量'
    );

    addSection(
      '👀 高留存但播放一般',
      COLORS.amber,
      signals.hidden,
      'retention_3s',
      '3秒留存'
    );

    addSection(
      '⚠️ 3秒留存较弱（本期后25%）',
      COLORS.red,
      signals.weak,
      'retention_3s',
      '3秒留存'
    );

    return sheet;
  }

  async function exportXlsx() {
    if (busy) {
      return;
    }

    const allRows = [
      ...detailRows
        .values()
    ]
      .sort(
        (a, b) =>
          b.publish_ts -
          a.publish_ts
      );

    const analysisRows =
      getAnalysisRows();

    if (
      !allRows.length
    ) {
      return;
    }

    if (
      typeof ExcelJS ===
      'undefined'
    ) {
      alert(
        'ExcelJS 组件未加载，请刷新页面后重试。'
      );

      return;
    }

    busy = true;

    $('qtk-progress-bar')
      .style
      .width =
      '0%';

    updateUI(
      '正在创建美化 Excel…'
    );

    try {
      const workbook =
        new ExcelJS
          .Workbook();

      workbook.creator =
        'Qualitell TikTok Studio Collector';

      workbook.created =
        new Date();

      workbook.modified =
        new Date();

      const imageIds =
        await prepareCoverImages(
          allRows,
          workbook
        );

      updateUI(
        '正在排版 Excel…'
      );

      $('qtk-progress-bar')
        .style
        .width =
        '100%';

      buildRawSheet(
        workbook,
        allRows,
        imageIds
      );

      buildAnalysisSheet(
        workbook,
        analysisRows,
        imageIds
      );

      const buffer =
        await workbook
          .xlsx
          .writeBuffer();

      const blob =
        new Blob(
          [
            buffer
          ],
          {
            type:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }
        );

      const account =
        allRows[0]
          ?.account ||
        'tiktok';

      const date =
        new Date()
          .toISOString()
          .slice(
            0,
            10
          );

      downloadBlob(
        blob,
        `${account}_TikTok周报_${date}.xlsx`
      );

      updateUI(
        `XLSX 已导出：${allRows.length} 条公开视频`
      );

    } catch (
      error
    ) {
      console.error(
        '[TikTok Collector] XLSX export failed',
        error
      );

      alert(
        'XLSX 导出失败，请打开控制台查看错误。'
      );

      updateUI(
        'XLSX 导出失败'
      );

    } finally {
      busy = false;

      setTimeout(
        () => {
          if (
            $('qtk-progress-bar')
          ) {
            $('qtk-progress-bar')
              .style
              .width =
              '0%';
          }

          updateUI();
        },
        800
      );
    }
  }

  function boot() {
    if (
      !document
        .documentElement
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
      document
        .addEventListener(
          'DOMContentLoaded',
          buildPanel,
          {
            once:
              true
          }
        );
    } else {
      buildPanel();
    }
  }

  boot();

})();
