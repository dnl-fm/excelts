
// =============================================================================
// AutoDrain - kind of /dev/null
import { EventEmitter } from 'events';

class AutoDrain extends EventEmitter {
  write(chunk) {
    this.emit('data', chunk);
  }

  end() {
    this.emit('end');
  }
}

export default AutoDrain;