declare namespace CustomFunctions {
  interface Invocation {
    address?: string;
    parameterAddresses?: string[];
  }

  interface CancelableInvocation extends Invocation {
    onCanceled: () => void;
  }

  interface StreamingInvocation<ResultType> extends CancelableInvocation {
    setResult(value: ResultType | Error): void;
  }

  function associate(id: string, func: (...args: any[]) => any): void;
}
