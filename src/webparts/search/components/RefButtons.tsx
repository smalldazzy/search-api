import * as React from 'react';
import { DefaultButton } from 'office-ui-fabric-react';

const RefButtons = (props) => (
    <div className='buttonslist'>
        <h1>Refiners</h1>
        {
            props.terms.map(item =>
                    <DefaultButton
                        data-automation-id="test"
                        text={item}
                        onClick={(e) => {
                            console.log(e.target);
                            props.search((e.target as HTMLElement).innerHTML);}}
                        allowDisabledFocus={true}
                    />
            )}
            <DefaultButton
                        data-automation-id="test"
                        text='Full data'
                        onClick={(e) => {
                            console.log(e.target);
                            props.search('');
                        }}
                        allowDisabledFocus={true}
                    />
    </div>
);
export default RefButtons;