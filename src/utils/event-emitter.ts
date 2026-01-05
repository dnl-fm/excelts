/**
 * Simple EventEmitter - Bun-native replacement for Node's events module
 */

type EventHandler = (...args: unknown[]) => void;

class SimpleEventEmitter {
  private handlers: Map<string, EventHandler[]> = new Map();
  private onceHandlers: Map<string, Set<EventHandler>> = new Map();

  on(event: string, handler: EventHandler): this {
    if (!this.handlers.has(event)) {
      this.handlers.set(event, []);
    }
    this.handlers.get(event)!.push(handler);
    return this;
  }

  once(event: string, handler: EventHandler): this {
    if (!this.onceHandlers.has(event)) {
      this.onceHandlers.set(event, new Set());
    }
    this.onceHandlers.get(event)!.add(handler);
    return this.on(event, handler);
  }

  off(event: string, handler: EventHandler): this {
    const handlers = this.handlers.get(event);
    if (handlers) {
      const index = handlers.indexOf(handler);
      if (index !== -1) {
        handlers.splice(index, 1);
      }
    }
    this.onceHandlers.get(event)?.delete(handler);
    return this;
  }

  removeListener(event: string, handler: EventHandler): this {
    return this.off(event, handler);
  }

  removeAllListeners(event?: string): this {
    if (event) {
      this.handlers.delete(event);
      this.onceHandlers.delete(event);
    } else {
      this.handlers.clear();
      this.onceHandlers.clear();
    }
    return this;
  }

  emit(event: string, ...args: unknown[]): boolean {
    const handlers = this.handlers.get(event);
    if (!handlers || handlers.length === 0) {
      return false;
    }

    const onceSet = this.onceHandlers.get(event);
    
    // Copy handlers to avoid mutation during iteration
    const handlersCopy = handlers.slice();
    for (const handler of handlersCopy) {
      handler(...args);
      
      // Remove if it was a once handler
      if (onceSet?.has(handler)) {
        this.off(event, handler);
      }
    }

    return true;
  }

  listenerCount(event: string): number {
    return this.handlers.get(event)?.length || 0;
  }

  listeners(event: string): EventHandler[] {
    return [...(this.handlers.get(event) || [])];
  }

  eventNames(): string[] {
    return [...this.handlers.keys()];
  }
}

export default SimpleEventEmitter;
export { SimpleEventEmitter, EventHandler };
