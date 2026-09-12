/* global Office, document, window */

(function () {
  var TYPE_LABEL = { to: '宛先', cc: 'CC', bcc: 'BCC' };
  var answered = false;

  function payload() {
    try {
      var raw = new URLSearchParams(window.location.search).get('p');
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }

  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function render(data) {
    var list = document.getElementById('list');
    var recipients = data.recipients || [];

    recipients.forEach(function (r) {
      var li = el('li', 'item');
      li.appendChild(el('div', 'addr', r.a || '(アドレス未解決)'));

      var meta = el('div', 'meta');
      meta.appendChild(document.createTextNode(TYPE_LABEL[r.t] || r.t || ''));
      if (r.n) meta.appendChild(document.createTextNode('・' + r.n));
      meta.appendChild(document.createTextNode('・'));
      meta.appendChild(
        el('span', 'domain', r.d ? '@' + r.d : 'ドメインを判定できません')
      );
      li.appendChild(meta);
      list.appendChild(li);
    });

    if (data.more > 0) {
      var more = document.getElementById('more');
      more.textContent = 'ほか ' + data.more + ' 件の宛先が許可リストに含まれていません。';
      more.hidden = false;
    }
  }

  function reply(action) {
    if (answered) return;
    answered = true;
    Office.context.ui.messageParent(JSON.stringify({ action: action }));
  }

  Office.onReady(function () {
    render(payload());

    document.getElementById('send').addEventListener('click', function () {
      reply('send');
    });
    document.getElementById('cancel').addEventListener('click', function () {
      reply('cancel');
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') reply('cancel');
    });

    document.getElementById('cancel').focus();
  });
})();
