# 🚀 업무 자동화 사례: 항공권 결제내역 정산(대조) 자동화 시스템

<br>

## 1. 프로젝트 배경 (Background)

피플팀에서 매월 수행하는 **항공권 정산 작업**은 반복적인 수작업과 피로도(특히 눈이 아픈..)를 동반하는 업무였습니다. 
항공권 법인카드 정산 회계전표의 첨부자료로 필수 반영되어야 하는 작업이지만, 
기존 수작업으로 진행하던 한계로 인해 담당자 개인의 시간과 리소스가 낭비되고 있었습니다.

<div class="claude-card">
  <h4>⚠️ 주요 페인 포인트 (Pain Point)</h4>
  <p>① 사내 e-DAHS 시스템과 세중여행사에서 제공하는 방대한 리스트를 모니터 양쪽에 띄워두고 <strong>일일이 눈으로 대조</strong>해야 했습니다. 
  
  ② 승인번호 중복, 누락, 또는 미세한 금액 불일치를 쉽게 발견하기가 어렵고, 발견하더라도 어디서 잘못되었는지 추적하기가 어려웠습니다.</p>
</div>

<div class="claude-card">
  <h4>🚨 비즈니스 리스크 (Risk)</h4>
  <p>① 수작업으로 인한 <strong>휴먼 에러 발생 가능성</strong>이 항상 존재합니다. 
  
  ② 1원 단위의 금액 오차라도 생길 경우, 회계전표는 상신 및 결재처리될 수 없습니다.
  
  ③ 또한, 법인카드 결제일 기한 내 완료하지 못하면 연체료가 부과되는 등의 리스크도 존재합니다.
  
  ④ 상기와 같은 사유로, 변수가 발생할 경우 담당자에겐 <strong>"처음부터 다시 대조해야하면 어쩌지?"</strong>라는 불안감이 항상 찾아옵니다. </p>
</div>

## 2. 해결 전략 (The Logic & Solution)

VLOOKUP과 같은 단순 엑셀 함수로는 예외 케이스가 워낙 많았습니다. 

이에 **Python 기반의 데이터 매칭 파이프라인**을 구축했습니다. v1부터 v9까지 로직을 고도화했으며, 그 중 4가지 핵심 개선 사항은 다음과 같습니다.

(파이썬 배경지식이 1도 없던 제가, **바이브코딩**의 도움을 받아 사용해보았습니다.)

<div class="claude-card logic-card">
  <h3>1. 🧹 데이터 정제 및 전처리 (Data Cleansing)</h3>
  <ul>
    <li><strong>승인번호 정규화</strong>: <code>001234</code>, <code> 1234 </code> 등 시스템마다 다르게 입력된 승인번호에서 숫자만 추출하고 앞의 '0'을 제거하여 고유 키(Key) 일관성 확보.</li>
    <li><strong>금액 데이터 정제</strong>: <code>KRW 1,000</code>, <code> 1,000.00 </code> 등 문자가 섞인 텍스트에서 불필요한 문자를 제거하고 순수 숫자(Float)로 변환.</li>
    <li><strong>공백 파괴</strong>: 이름 데이터 띄어쓰기를 전부 제거하여 <code>HONG GILDONG</code>과 <code>HONGGILDONG</code>을 동일 인물로 정확히 매칭.</li>
  </ul>
</div>

<div class="claude-card logic-card">
  <h3>2. 🔑 다차원 승인번호 매칭 (Multi-Key Matching)</h3>
  <ul>
    <li>단순 VLOOKUP의 한계를 넘어 <strong>두 개의 식별자</strong>를 동시에 검증.</li>
    <li>세중여행사 리스트에서 건당 매칭 시, <strong>승인번호(AC열)</strong>와 <strong>취급승인번호(AV열)</strong>를 동시에 확인하여 어느 한쪽이라도 일치하면(OR 조건) 통과시키는 입체적 검사 로직 적용.</li>
  </ul>
</div>

<div class="claude-card logic-card">
  <h3>3. ⚖️ 입체적 금액 검증 (Amount Validation)</h3>
  <ul>
    <li><strong>항공권 가격</strong>: 승인번호 매칭 시 세중여행사의 <code>공급가액</code>과 e-DAHS의 <code>총합계</code> 대조.</li>
    <li><strong>취급 수수료</strong>: 취급승인번호 매칭 시 세중여행사의 <code>취급수수료</code>와 e-DAHS의 <code>총합계</code> 대조.</li>
    <li><strong>정밀 오차 감지</strong>: 결제내역 누락, 환불 패널티 등으로 인한 단 <strong>1원의 오차</strong>도 에러(Red)로 판정하는 100% 무결성 검증.</li>
  </ul>
</div>


## 3. 핵심 기능 및 시각화 피드백 (Key Features)

실무자가 엑셀을 열었을 때 **직관적으로 문제 상황만 파악**할 수 있도록 파일에 직접 색상을 칠하는 자동 하이라이트 기능을 탑재했습니다.

<div class="claude-alert error-alert">
  <strong>🔴 빨간색 (Error - 금액 불일치)</strong>
  <p>승인번호는 일치하나 금액이 단 1원이라도 불일치하는 경우. 결제 프로세스 오류나 환율 적용 차이를 의미하며, 즉각적인 원인 파악 및 재검토가 가능해집니다.</p>
</div>

<div class="claude-alert warning-alert">
  <strong>🟡 노란색 (Missing - 누락 건)</strong>
  <p>세중여행사 리스트에는 있으나 e-DAHS 결제 내역에는 없는 경우. 누락된 영수증이나 미처리 결제 건을 식별하여 정산 누락을 원천 차단합니다.</p>
</div>

<div class="claude-card success-card">
  <h4>✨ 핵심 기능: 자동 기입 (Auto-fill)</h4>
  <p>세중여행사 리스트에만 존재하는 정보(탑승자 영문명, 출발일, 여정)를 e-DAHS 결제내역 시트의 알맞은 위치에 <strong>자동으로 꽂아넣어</strong> 수작업(복사-붙여넣기)을 완전히 없앴습니다.</p>
</div>

## 4. 🔄 업무 환경의 변화 (Before & After)

| 구분 (Category) | 📉 Before (도입 이전) | 🚀 After (자동화 이후) |
| :--- | :--- | :--- |
| **소요 시간** | 최소 2~3시간 걸리던 대조 작업 | 단 **수 초(Seconds)** 만에 100% 완료 |
| **작업 방식** | 듀얼 모니터로 최소 수백 건의 엑셀 데이터를 **일일이 눈으로 대조** | 파이썬 파이프라인이 자동 검증 후 **결과 시트 즉시 생성** |
| **정확도** | 휴먼 에러 위험 상존, 1원 단위 오차 발견 어려움 | 1원의 오차도 놓치지 않는 **100% 무결성 검증** |
| **담당자 역할** | 단순 반복적인 '틀린 그림 찾기' (극심한 피로도) | 에러 건(빨간색/노란색 셀) **원인 분석 및 조치**에만 집중 |
| **유지보수** | 각기 다른 포맷으로 관리의 어려움 | 데이터 양식에 변화가 생기더라도 **바이브코딩**을 통해 수정만 해주면 완료 |

<br>

<div class="claude-card">
  <h3>🎥 Before: 기존의 수작업 대조</h3>
  <p>자동화 이전에는 아래 영상처럼 듀얼 모니터에 두 개의 엑셀을 띄워두고, 수천 건의 데이터를 <strong>일일이 눈으로 대조</strong>해야 했습니다.<br>
  시간이 오래 걸리고 눈이 피로해지는 것은 물론, 집중력이 떨어지면 미세한 오차나 누락 건을 놓치기 일쑤였습니다. (다른 업무에서도 집중력 ↓)</p>
  <!-- 👇 아래 src 속성에 수작업 녹화 영상(mp4) 파일 경로를 지정해 주세요 -->
  <video src="media/before_manual.mp4?v=2" controls style="width: 100%; border-radius: 8px; margin-top: 12px; border: 1px solid rgba(255,255,255,0.1); background: #000;"></video>
</div>

<div class="claude-card success-card">
  <h3>⚡ After: 원클릭 자동화 및 결과 즉시 도출</h3>
  <p>자동화 도입 후, <strong>단 한 번의 파이썬 스크립트 실행(One-click)</strong>만으로 몇시간이 걸리던 대조 작업이 수 초 만에 끝납니다. </p>
  <!-- 👇 아래 src 속성에 파이썬 자동화 구동 영상(mp4) 파일 경로를 지정해 주세요 -->
  <video src="media/after_manual.mp4?v=1" controls style="width: 100%; border-radius: 8px; margin-top: 12px; border: 1px solid rgba(255,255,255,0.1); background: #000;"></video>
</div>

<br>

---

**💡 결론 및 인사이트 (Insights)**  
항공권 정산업무와 같은 반복적인 대사 업무는 어느 부서에나 존재한다고 생각합니다. 

이번 사례처럼 <strong>"명확한 식별자(Key)"</strong>와 <strong>"검증 규칙"</strong>만 정의된다면, 수작업의 늪에서 벗어나 엄청난 효율성을 높일 수 있습니다. 

물론 처음 세팅하는 과정이 막막하고 힘들 수 있지만, AI 툴에게 **내가 했던 방식을 차근차근 설명**해주면 이후엔 편하게 산출물을 얻어낼 수 있습니다. **(🍺고진감래🍺)**
