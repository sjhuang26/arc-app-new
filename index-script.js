$(document).ready(runApp);
function runApp() {
  // UTILITIES
  class Event {
    constructor() {
      this.listeners = [];
    }
    trigger() {
      for (const listener of this.listeners) {
        listener();
      }
    }
    get chain() {
      return this.trigger.bind(this);
    }
    listen(callback) {
      this.listeners.push(callback);
    }
  }
  class MutableValue {
    constructor(initialValue) {
      this.val = initialValue;
      this.changeEvent = new Event();
      onMount.listen(this.changeEvent.chain);
    }
  }
  
  // GLOBAL EVENTS
  const onMount = new Event();

  // ROOT WIDGET
  $('#app').append(makeRootWidget());
  onMount.trigger();

  function makeRootWidget() {
    // UTILITIES
    function container(newTag) {
      return (...children) => $(newTag).append(...children);
    }

    // STATE
    let pageState = new MutableValue(1);

    // DOM
    const $dom = container('<div class="container-fluid"></div>')(
      makeNavWidget(),
      container('<div class="row"></div>')(
        container('<div class="col"></div>')(
        )
      )
    );

    // CHILDREN
    function makeSpinnerWidget(isLoadedState, $wrapped) {
      isLoadedState.changeEvent.listen(() => {
        if (isLoadedState.val) {
          $spinner.hide();
          $wrapped.show();
        } else {
          $spinner.show();
          $wrapped.hide();
        }
      });

      const $spinner = container('<div></div>')(
        $('<strong>Loading...</strong>'),
        $('<div class="spinner-border"></div>')
      );

      const $dom = container('<div></div>')(
        $spinner,
        $wrapped
      );

      return $dom;
    }
    function makeButtonWidget(text, onClick, variant) {
      // to create an outline button, add "outline" to the variant
      variant = variant || 'primary';
      if (variant === 'outline') variant = 'outline-primary';
      return $('<button></button>')
        .text(text)
        .addClass('btn btn-' + variant)
        .click(onClick);
    }
    function makeTabsWidget(state, values) {
      state.changeEvent.listen(() => {
        $dom.children().find('a').removeClass('active');
        $dom.children().eq(state.val).find('a').addClass('active');
      });

      function makeNavItemWidget(text, pageNumber) {
        const link = $('<a class="nav-link"></a>').text(text);
        return $('<li class="nav-item"></li>').append(link).click(() => {
          state.val = pageNumber;
          state.changeEvent.trigger();
        });
      }

      const $dom = container('<ul class="nav nav-tabs"></ul>')(
        values.map((value, index) => makeNavItemWidget(value, index))
      );
      return $dom;
    }
    function makeNavWidget() {
      return makeTabsWidget(pageState, ['Apple', 'Banana', 'Cherry']);
    }

    return $dom;
  }
}
